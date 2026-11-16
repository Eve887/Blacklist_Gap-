"""
build_network_firms_buyers.py - the start of the network layer: one row per
firm, one column per buyer, confirmed EXACT cross-buyer matches only.

Not part of the numbered pipeline (not auto-chained) - run it by hand
whenever cross_buyer_cases.xlsx / master_list.xlsx change.

Reads:
  output/master_list.xlsx   ("Master list" + "Matches" sheets)
Writes:
  output/network_firms_buyers.xlsx

Only uses the Matches sheet's confirmed EXACT standard-name matches across
DIFFERENT buyers (NOT the fuzzy Part C candidates in match_candidates.xlsx -
those are unreviewed guesses, not confirmed matches, and do not belong in
a "confirmed network" file). Each cell holds every date range that buyer
has for that firm (semicolon-separated if more than one), blank if that
buyer has no entry for this firm at all.
"""

from collections import defaultdict

from openpyxl import load_workbook, Workbook

from common import OUT, write_sheet


def main():
    wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    ws = wb["Master list"]
    rows = list(ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    id_i = header.index("Entry ID")
    buyer_i = header.index("Buyer")
    from_i = header.index("Blacklisted from")
    to_i = header.index("Blacklisted to")
    by_id = {r[id_i]: r for r in data}

    all_buyers = sorted({r[buyer_i] for r in data})

    match_rows = list(wb["Matches"].iter_rows(values_only=True))[1:]

    network = {}                                            # firm -> {buyer: [date ranges]}
    for firm, times_listed, entry_ids_text, buyers_text, meaning in match_rows:
        if not meaning or not str(meaning).startswith("DIFFERENT"):
            continue
        entry_ids = [e.strip() for e in entry_ids_text.split(",")]
        cell = network.setdefault(firm, defaultdict(list))
        for eid in entry_ids:
            r = by_id.get(eid)
            if r is None:
                continue
            cell[r[buyer_i]].append(f"{r[from_i]} to {r[to_i]}")

    out_rows = []
    for firm in sorted(network):
        row = [firm]
        for buyer in all_buyers:
            dates = network[firm].get(buyer)
            row.append("; ".join(sorted(set(dates))) if dates else "")
        out_rows.append(row)

    header_out = ["Firm (standard name)"] + all_buyers
    widths = [36] + [26] * len(all_buyers)

    wb_out = Workbook()
    write_sheet(wb_out.active, header_out, out_rows, widths)
    wb_out.active.title = "Network - firms x buyers"
    out = OUT / "network_firms_buyers.xlsx"
    wb_out.save(out)

    print(f"Saved: {out}")
    print(f"  Firms (confirmed exact cross-buyer matches): {len(out_rows)}")
    print(f"  Buyer columns: {all_buyers}")


if __name__ == "__main__":
    main()
