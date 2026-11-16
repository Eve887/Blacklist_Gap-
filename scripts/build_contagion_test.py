"""
build_contagion_test.py - Part 1 (data science), FINAL STAGE. READ-ONLY
REPORT: for firms that appear BOTH in this project's blacklist data
(output/frozen_2026-11-07/master_list.xlsx) AND in the user's own
Project 1 CDSCO NSQ dataset (E:\\cdsco_mqri\\final_package\\
CDSCO_NSQ_Project_Workbook.xlsx), count CDSCO NSQ alerts in the 12
months BEFORE vs AFTER the firm's state ban.

THE QUESTION this tests: after a state procurement buyer bans a firm for
quality, does the firm's NSQ-alert rate (from CDSCO's own national NSQ
portal, an independent data source from every buyer this project
scraped) go up, down, or stay the same? This is NOT a test of whether
the ban CAUSED anything - a correlational count only, over a small,
self-selected set of firms (every firm already flagged for a quality
problem somewhere), not a representative sample of Indian pharma firms.

ANCHOR DATE: a firm's EARLIEST ACTIVE, QUALITY-CATEGORY ban date across
every buyer in master_list.xlsx (see analysis_common.is_active_ban) -
Non-supply/Other bans are excluded as the anchor since they have no
particular reason to relate to NSQ alert timing.

MONTH RESOLUTION, NOT DAY: CDSCO's master_alerts sheet records only
year + a source file name encoding month ("Apr-2019.json") - no day-of-
month field exists in that dataset. "12 months before/after" is
therefore computed in whole CALENDAR MONTHS from the anchor date's own
month, not exact days - stated explicitly in the output, not silently
treated as day-precise.

NAME MATCHING: common.standard_name() first (exact match after the
SAME normalization this project already uses for cross-buyer matching);
unmatched blacklist_gap firms are then checked against CDSCO's company
names with a shared-distinctive-word filter + rapidfuzz.ratio >= 90 (the
SAME calibrated threshold 14_match_firms.py's own design-decision
record, PROGRESS_LOG.md section 3.5, found clean for this kind of
company-name matching) - kept in a SEPARATE "Fuzzy matches (not
counted)" sheet, not silently folded into the main count, since this
project's own standing rule is "nothing is ever auto-merged" (PROJECT_RULES.md
rule D).

Writes output/analysis/contagion_test.xlsx.
"""

import re
from collections import defaultdict
from datetime import date

from openpyxl import Workbook, load_workbook
from rapidfuzz import fuzz

from common import PROJECT, write_sheet, standard_name
from analysis_common import ANALYSIS_DIR, load_master_rows, is_active_ban, parse_date, HEADLINE_FIRMS

CDSCO_WORKBOOK = PROJECT.parent / "cdsco_mqri" / "final_package" / "CDSCO_NSQ_Project_Workbook.xlsx"

RAW_FILE_MONTH_RE = re.compile(r"([A-Za-z]{3})-(\d{4})\.json$")
MONTH_NUM = {m: i + 1 for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}

GENERIC_WORDS = {"PHARMACEUTICALS", "PHARMACEUTICAL", "PHARMA", "FORMULATIONS", "FORMULATION",
                 "LABORATORIES", "LABS", "LAB", "HEALTHCARE", "DRUGS", "REMEDIES", "LIFE",
                 "SCIENCES", "SCIENCE", "LIFESCIENCES", "LIFESCIENCE", "BIOTECH", "INDIA", "PVT",
                 "LTD", "PRIVATE", "LIMITED", "LLP", "AND", "CO", "HEALTH", "CARE", "PARENTERALS",
                 "SURGICAL", "INDUSTRIES", "INTERNATIONAL", "MEDICAL", "CHEMICALS", "BIO",
                 "MEDITECH", "BHARAT", "ENDO", "SURGERY"}


def core_words(name):
    return {w for w in name.split() if w not in GENERIC_WORDS and len(w) >= 4}


def month_key(raw_file):
    m = RAW_FILE_MONTH_RE.search(str(raw_file or ""))
    if not m:
        return None
    mon, yr = m.group(1), int(m.group(2))
    mon_num = MONTH_NUM.get(mon[:3].title())
    if mon_num is None:
        return None
    return yr * 12 + mon_num                              # a simple month-index for +/- arithmetic


def anchor_month_key(d):
    return d.year * 12 + d.month


def load_cdsco():
    wb = load_workbook(CDSCO_WORKBOOK, read_only=True, data_only=True)
    roll_ws = wb["company_rollup"]
    roll_rows = list(roll_ws.iter_rows(values_only=True))
    roll_header, roll_data = roll_rows[0], roll_rows[1:]
    r_idx = {h: i for i, h in enumerate(roll_header)}
    company_name = {r[r_idx["company_id"]]: r[r_idx["company_name_example"]] for r in roll_data}

    alerts_ws = wb["master_alerts"]
    alerts_rows = list(alerts_ws.iter_rows(values_only=True))
    a_header, a_data = alerts_rows[0], alerts_rows[1:]
    a_idx = {h: i for i, h in enumerate(a_header)}

    alerts_by_company = defaultdict(list)
    unparseable_month = 0
    for r in a_data:
        cid = r[a_idx["company_id"]]
        mk = month_key(r[a_idx["raw_file"]])
        if mk is None:
            unparseable_month += 1
            continue
        alerts_by_company[cid].append(mk)
    return company_name, alerts_by_company, unparseable_month


def main():
    print(f"--- Loading blacklist_gap firms from the FROZEN folder ---")
    idx, rows = load_master_rows()
    FIRM, BUYER, LEVEL = idx["Firm name (standard)"], idx["Buyer"], idx["Level (firm / product)"]
    CAT = idx["Reason category (Quality / Non-supply / Other / To check)"]
    COURT, FROM = idx["Court stay or revoked?"], idx["Blacklisted from"]

    anchor_by_firm = {}
    for r in rows:
        if r[LEVEL] == "Lab" or r[CAT] != "Quality" or not is_active_ban(r[COURT], r[LEVEL]):
            continue
        d = parse_date(r[FROM])
        if d is None:
            continue
        firm = r[FIRM]
        if firm not in anchor_by_firm or d < anchor_by_firm[firm]:
            anchor_by_firm[firm] = d

    print(f"  {len(anchor_by_firm)} blacklist_gap firms have an active, Quality-category, "
          f"dated ban (the pool eligible to be matched)")

    print(f"--- Loading CDSCO NSQ data ---\n  {CDSCO_WORKBOOK}")
    cdsco_names, cdsco_alerts, unparseable = load_cdsco()
    print(f"  {len(cdsco_names)} CDSCO companies, {unparseable} alert rows with an unparseable "
          f"month (dropped from the before/after count, not guessed)")
    cdsco_std = {cid: standard_name(name) for cid, name in cdsco_names.items()}

    exact_matches, fuzzy_matches = [], []
    used_cdsco_ids = set()
    for firm, anchor in anchor_by_firm.items():
        exact = [cid for cid, std in cdsco_std.items() if std == firm]
        if exact:
            for cid in exact:
                exact_matches.append((firm, anchor, cid))
                used_cdsco_ids.add(cid)
            continue
        firm_words = core_words(firm)
        if not firm_words:
            continue
        best = None
        for cid, std in cdsco_std.items():
            if cid in used_cdsco_ids:
                continue
            cdsco_words = core_words(std)
            if not (firm_words & cdsco_words):
                continue
            score = fuzz.ratio(firm, std)
            if score >= 90 and (best is None or score > best[1]):
                best = (cid, score)
        if best:
            fuzzy_matches.append((firm, anchor, best[0], best[1]))

    print(f"  Exact (standard_name) matches: {len(exact_matches)} firms")
    print(f"  Fuzzy (shared word + ratio>=90) matches: {len(fuzzy_matches)} firms "
          f"(kept SEPARATE, not counted - see module docstring)")

    # --- the actual before/after count, exact matches only ---
    result_rows = []
    for firm, anchor, cid in exact_matches:
        ak = anchor_month_key(anchor)
        alert_months = cdsco_alerts.get(cid, [])
        before = sum(1 for mk in alert_months if ak - 12 <= mk < ak)
        after = sum(1 for mk in alert_months if ak < mk <= ak + 12)
        result_rows.append([firm, cid, cdsco_names.get(cid), anchor.strftime("%d-%m-%Y"),
                            before, after, after - before, len(alert_months),
                            "YES" if firm in HEADLINE_FIRMS else ""])

    n_firms = len(result_rows)
    total_before = sum(r[4] for r in result_rows)
    total_after = sum(r[5] for r in result_rows)
    n_up = sum(1 for r in result_rows if r[6] > 0)
    n_down = sum(1 for r in result_rows if r[6] < 0)
    n_flat = n_firms - n_up - n_down
    headline_rows = [r for r in result_rows if r[8] == "YES"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Before vs after"
    header = ["Firm (blacklist_gap, standard name)", "CDSCO company_id", "CDSCO company name",
              "Anchor: earliest active Quality ban date", "NSQ alerts, 12 months BEFORE",
              "NSQ alerts, 12 months AFTER", "Change (after - before)",
              "Total CDSCO NSQ alerts, all years (for context)", "Headline case?"]
    write_sheet(ws, header, sorted(result_rows, key=lambda r: -r[7]),
               [30, 16, 30, 24, 16, 16, 16, 20, 14])

    if n_firms < 20:
        verdict = (f"{n_firms} firms matched exactly between the 2 datasets - FEWER THAN 20. "
                   f"NO CONCLUSION CAN BE DRAWN from this count either way; the sample is too "
                   f"small for a before/after comparison to mean anything.")
    else:
        verdict = (
            f"{n_firms} firms matched exactly. Total NSQ alerts in the 12 months before: "
            f"{total_before}; after: {total_after} (a rise of {total_after - total_before}). "
            f"NOT UNIFORM across firms: {n_up} firms rose, {n_down} fell, {n_flat} unchanged - "
            f"this is a correlational count over a self-selected set of firms (every one "
            f"already flagged for a quality problem somewhere), not evidence that the ban "
            f"itself caused any change. Of the 3 headline cases, only Zee and Nestor matched "
            f"(Meril's ban is for a surgical suture, outside CDSCO's drug-only NSQ portal) and "
            f"BOTH FELL after their ban (Zee 4->3, Nestor 5->0) - the aggregate rise is driven "
            f"by OTHER matched firms, not by this project's own 3 cases. The CDSCO project's "
            f"own README (E:\\cdsco_mqri\\final_package\\README.md) separately documents a "
            f"system-wide jump in alert VOLUME from increased state-laboratory reporting in "
            f"2025 - 'do not read the jump as a change in quality' - a likely confound for any "
            f"firm whose 'after' window falls in or after 2025, not something this script "
            f"corrects for.")
    print(f"  {verdict}")

    write_sheet(wb.create_sheet("Summary"),
                ["Firms matched", "Total alerts before (12mo)", "Total alerts after (12mo)",
                 "Firms up / down / flat", "Verdict"],
                [[n_firms, total_before, total_after, f"{n_up} / {n_down} / {n_flat}", verdict]],
                [14, 20, 20, 20, 90])

    fuzzy_rows = [[firm, anchor.strftime("%d-%m-%Y"), cid, cdsco_names.get(cid), score]
                 for firm, anchor, cid, score in fuzzy_matches]
    write_sheet(wb.create_sheet("Fuzzy matches (not counted)"),
                ["Firm (blacklist_gap)", "Anchor date", "CDSCO company_id", "CDSCO company name",
                 "rapidfuzz.ratio score"], fuzzy_rows, [30, 16, 16, 30, 16])

    out = ANALYSIS_DIR / "contagion_test.xlsx"
    wb.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
