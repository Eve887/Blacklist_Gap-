"""
build_rmsc_rc_wayback.py - Step 4b history: one Wayback Machine snapshot
of RMSC's DrugsRC.html per year (closest to 1 July), parsed with the SAME
code as 15_rate_contracts.py, then the Group A/B/C/D/X check re-run for
each year.

Not part of the numbered pipeline - a one-off report script, run by
hand. Confirmed via a direct CDX query (see PROGRESS_LOG.md Entry 7):
69 distinct-day snapshots exist, 2018-05-20 to 2026-04-20.

Important limitation, stated here rather than hidden: the Group A/B/C/D/X
check below cross-references each YEAR's historical RC data against
master_list.xlsx's CURRENT ban records - there is no historical snapshot
of every buyer's blacklist page to check against for each past year, so
this answers "given what we know about bans NOW, did this year's RC
suppliers overlap with a ban", not "what did the picture look like using
only information available at that time". Both are useful, but they are
not the same question.

Two more limitations found while building this, both logged to
output/source_issues.csv rather than silently worked around:

1. RMSC's table SCHEMA changed over time (confirmed by reading each
   snapshot's own header row, not assumed): 2018-2019 captures have only
   9 columns (no Tender No, no Level at all); 2020-2021 have 10 (added
   NIB/Tender No, still no Level); 2022 onward have the current 11
   (added Level). 15_rate_contracts.py's own parser is correctly scoped
   to the CURRENT 11-column page and rejects anything shorter (by
   design, to avoid misassigning a GST-rate value into a Level column,
   say) - so this script has its OWN schema-aware parser,
   `parse_rows_historical()`, instead of reusing that one directly for
   the 9/10-column years. Tender No and Level are honestly "Not stated"
   for 2018-2021, since those fields genuinely did not exist on the page
   yet - not guessed at.
2. The 2023-09-26 and 2024-06-25 captures are themselves TRUNCATED at
   the Internet Archive - both end abruptly mid-`<tr>` tag with no
   closing `</table>`/`</html>`, confirmed by requesting the SAME
   snapshot both with and without the "id_" raw-content modifier (same
   cut-off point either way, so this is IA's own stored capture, not
   something this script's request is doing). This means the 2023/2024
   row counts are a floor, not a complete count, for those two years
   only.

Download pattern: web.archive.org/web/<timestamp>id_/<url> - the "id_"
suffix returns the raw captured HTML without Wayback's toolbar/banner
injection, same convention used elsewhere for Wayback captures in this
project. 5-second pauses between downloads, SHA-256 logged to the shared
output/download_log.csv.

Writes:
  data/raw/RMSC_RC/wayback/DrugsRC_<snapshot date>.html  - one per year
  output/rate_contracts/RMSC_history.xlsx  - "All years" (RC rows, Date
    captured = snapshot date, not today) + "By year summary" (A-X firm
    counts per year, same classify() logic as build_rc_vs_bans.py)
"""

import csv
import time
import hashlib
import importlib.util
from datetime import date, datetime
from collections import defaultdict
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

from common import OUT, RAW, write_sheet, dedup_entry_ids

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 5
TIMEOUT = 60
ORIGINAL_URL = ("https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-"
                "services-corporation-ltd-/en/Approved-Rate-Lists/DrugsRC.html")

WAYBACK_DIR = RAW / "RMSC_RC" / "wayback"
WAYBACK_DIR.mkdir(parents=True, exist_ok=True)

# 15_rate_contracts.py and build_rc_vs_bans.py both have filenames that
# can't be imported with a normal "import" statement (one starts with a
# digit) - loaded by file path, same pattern as 06_gmscl_archive_playwright.py.
_rc_spec = importlib.util.spec_from_file_location("rc_module", Path(__file__).parent / "15_rate_contracts.py")
rc_module = importlib.util.module_from_spec(_rc_spec)
_rc_spec.loader.exec_module(rc_module)

_bans_spec = importlib.util.spec_from_file_location("bans_module", Path(__file__).parent / "build_rc_vs_bans.py")
bans_module = importlib.util.module_from_spec(_bans_spec)
_bans_spec.loader.exec_module(bans_module)

session = requests.Session()
session.headers.update(HEADERS)


def parse_rows_historical(soup):
    """Schema-aware parser for Wayback snapshots - handles the 9-column
    (2018-2019), 10-column (2020-2021) and current 11-column (2022+)
    table layouts, confirmed by reading each snapshot's own header row
    (see the module docstring). Returns the same entry-dict shape
    15_rate_contracts.py's build_row() expects."""
    table = soup.find("table")
    if table is None:
        return [], 0

    rows = table.find_all("tr")
    header_cols = len(rows[0].find_all(["td", "th"])) if rows else 0
    out = []
    for tr in rows:
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["td", "th"])]
        if not cells or cells[0].strip().upper() == "DRUG NAME":
            continue
        if header_cols == 11 and len(cells) >= 11:
            (drug_name, code, tender_no, supplier, _contact, unit, rate, gst, level,
             frm, to) = cells[:11]
        elif header_cols == 10 and len(cells) >= 10:
            (drug_name, code, tender_no, supplier, _contact, unit, rate, gst,
             frm, to) = cells[:10]
            level = "Not stated"
        elif header_cols == 9 and len(cells) >= 9:
            (drug_name, code, supplier, _contact, unit, rate, gst, frm, to) = cells[:9]
            tender_no, level = "Not stated", "Not stated"
        else:
            continue
        out.append({"drug_name": drug_name, "code": code, "tender_no": tender_no,
                    "supplier": supplier, "unit": unit, "rate": rate, "gst": gst,
                    "level": level, "frm": frm, "to": to})
    return out, header_cols


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def log_source_issue(year, snapshot_date, html_path, header_cols):
    """Appends to output/source_issues.csv using its REAL schema
    (date,buyer,file,page,issue,status) - a wrong schema was used by
    mistake once already in this project (see PROGRESS_LOG.md Entry 8),
    so this is written out explicitly rather than guessed again."""
    path = OUT / "source_issues.csv"
    is_new = not path.exists()
    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["date", "buyer", "file", "page", "issue", "status"])
        if is_new:
            w.writeheader()
        w.writerow({
            "date": date.today().strftime("%d-%m-%Y"), "buyer": "RMSC",
            "file": html_path.name, "page": "-",
            "issue": f"Wayback snapshot for {year} ({snapshot_date}), {header_cols}-column "
                     f"schema, is TRUNCATED at the Internet Archive itself - the HTML ends "
                     f"mid-<tr> with no closing </table>/</html>. Confirmed this is IA's own "
                     f"stored capture, not a download issue here (same cut-off point requesting "
                     f"the snapshot both with and without the 'id_' raw-content modifier).",
            "status": f"Open - row count for {year} in RMSC_history.xlsx is a FLOOR, not "
                       f"complete. No full, untruncated copy of this snapshot is known to exist.",
        })


def log_download(status, file_path, sha256, url):
    log_path = OUT / "download_log.csv"
    is_new = not log_path.exists()
    with open(log_path, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["date_time", "status", "file", "sha256", "url"])
        if is_new:
            w.writeheader()
        w.writerow({"date_time": datetime.now().strftime("%d-%m-%Y %H:%M"), "status": status,
                    "file": str(file_path), "sha256": sha256, "url": url})


def fetch_cdx_snapshots():
    """Live CDX query (not hardcoded) - returns a sorted list of
    (timestamp_str, year) for every distinct-day snapshot found."""
    r = requests.get("https://web.archive.org/cdx/search/cdx", params={
        "url": ORIGINAL_URL.replace("https://", "").replace("http://", ""),
        "output": "json", "matchType": "prefix", "collapse": "timestamp:8", "limit": "10000",
    }, timeout=60)
    r.raise_for_status()
    import json
    rows = json.loads(r.text)
    out = []
    for row in rows[1:]:
        ts = row[1]
        out.append((ts, int(ts[:4])))
    return sorted(out)


def pick_closest_to_july1(snapshots):
    """One per year - the snapshot whose date is closest to 1 July that year."""
    by_year = defaultdict(list)
    for ts, year in snapshots:
        by_year[year].append(ts)

    chosen = {}
    for year, timestamps in by_year.items():
        july1 = date(year, 7, 1)

        def distance(ts):
            d = datetime.strptime(ts[:8], "%Y%m%d").date()
            return abs((d - july1).days)

        chosen[year] = min(timestamps, key=distance)
    return chosen


def main():
    print("--- Querying Wayback CDX for all DrugsRC.html snapshots ---")
    snapshots = fetch_cdx_snapshots()
    print(f"Found {len(snapshots)} distinct-day snapshots, "
          f"{snapshots[0][0][:8]} to {snapshots[-1][0][:8]}.")

    chosen = pick_closest_to_july1(snapshots)
    print(f"Picked {len(chosen)} snapshots (one per year, closest to 1 July):")
    for year in sorted(chosen):
        print(f"  {year}: {chosen[year]}")

    all_rows = []
    by_year_entries = {}   # year -> list of build_row() output, for the classify step
    year_meta = {}         # year -> (header_cols, truncated)
    for year in sorted(chosen):
        ts = chosen[year]
        snapshot_date = datetime.strptime(ts[:8], "%Y%m%d").strftime("%Y-%m-%d")
        wayback_url = f"https://web.archive.org/web/{ts}id_/{ORIGINAL_URL}"
        print(f"\n--- {year}: fetching snapshot {snapshot_date} ---")

        try:
            r = session.get(wayback_url, timeout=TIMEOUT)
            r.raise_for_status()
        except requests.RequestException as e:
            print(f"  DOWNLOAD FAILED: {e} - skipping this year.")
            time.sleep(WAIT_SECONDS)
            continue

        html_path = WAYBACK_DIR / f"DrugsRC_{snapshot_date}.html"
        html_path.write_text(r.text, encoding="utf-8")
        sha = sha256_bytes(r.content)
        log_download("captured", html_path, sha, wayback_url)
        print(f"  Saved: {html_path}  ({len(r.content):,} bytes)")
        print(f"  SHA-256 logged to output/download_log.csv: {sha}")

        truncated = "</table>" not in r.text.lower() or "</html>" not in r.text.lower()
        if truncated:
            print(f"  WARNING: this capture is TRUNCATED (no closing </table>/</html> found) - "
                  f"the Internet Archive's own stored snapshot ends mid-row. Row count below is "
                  f"a FLOOR for this year, not complete. Logged to source_issues.csv.")

        soup = BeautifulSoup(r.text, "html.parser")
        entries, header_cols = parse_rows_historical(soup)
        rows = []
        for e in entries:
            row = rc_module.build_row(e)
            row[12] = snapshot_date          # "Date captured" = snapshot date, not today
            row[13] = wayback_url            # "Source link" = the actual evidence location
            rows.append(row)
        dupe_warnings = dedup_entry_ids(rows)
        if dupe_warnings:
            print(f"  WARNING - duplicate Entry IDs in {year} snapshot:")
            for w in dupe_warnings:
                print("   ", w)

        print(f"  Table schema: {header_cols} columns. Parsed {len(rows)} rate-contract rows.")
        for row in rows:
            all_rows.append([year, snapshot_date] + row)
        by_year_entries[year] = rows
        year_meta[year] = (header_cols, truncated)
        if truncated:
            log_source_issue(year, snapshot_date, html_path, header_cols)
        time.sleep(WAIT_SECONDS)

    # Re-run the Group A/B/C/D/X check for each year, against the CURRENT
    # master_list.xlsx (see the module docstring for why this is a
    # "bans as known now" check, not a fully historical one).
    master_wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    master_ws = master_wb["Master list"]
    master_rows = list(master_ws.iter_rows(values_only=True))
    m_header, m_data = master_rows[0], master_rows[1:]
    m_idx = {name: i for i, name in enumerate(m_header)}
    bans_by_name = defaultdict(list)
    for r in m_data:
        if r[m_idx["Buyer"]] == "RMSC":
            continue
        bans_by_name[r[m_idx["Firm name (standard)"]]].append(r)

    rc_header = ["Entry ID", "Drug Name", "Drug Code", "Tender No", "Supplier (as written)",
                "Supplier (standard)", "Manufacturer (if stated)", "Rate", "GST Rate", "Level",
                "Contract From", "Contract To", "Date captured", "Source link"]
    rc_idx = {name: i for i, name in enumerate(rc_header)}

    summary_rows = []
    for year in sorted(by_year_entries):
        firms_by_group = defaultdict(set)
        for rc in by_year_entries[year]:
            supplier_std = rc[rc_idx["Supplier (standard)"]]
            bans = bans_by_name.get(supplier_std)
            if not bans:
                continue
            rc_from = bans_module.parse_date(rc[rc_idx["Contract From"]])
            rc_to = bans_module.parse_date(rc[rc_idx["Contract To"]])
            for ban in bans:
                ban_from = bans_module.parse_date(ban[m_idx["Blacklisted from"]])
                ban_to = bans_module.parse_date(ban[m_idx["Blacklisted to"]])
                ov = bans_module.overlap(rc_from, rc_to, ban_from, ban_to)
                if ov != "yes":
                    continue
                group, _why = bans_module.classify(
                    ban[m_idx["Level (firm / product)"]],
                    ban[m_idx["Reason category (Quality / Non-supply / Other / To check)"]],
                    ban[m_idx["Court stay or revoked?"]],
                    rc[rc_idx["Drug Name"]], ban[m_idx["Product(s)"]])
                firms_by_group[group].add(supplier_std)
        header_cols, truncated = year_meta[year]
        row = [year, chosen[year], header_cols, "Yes" if truncated else "No",
               len(by_year_entries[year])]
        for g in "ABCDUX":
            row.append(len(firms_by_group.get(g, set())))
        summary_rows.append(row)
        print(f"\n{year} ({chosen[year]}, {header_cols}-col schema"
              f"{', TRUNCATED' if truncated else ''}): {len(by_year_entries[year])} RC rows -> "
              f"A={len(firms_by_group.get('A', set()))} "
              f"B={len(firms_by_group.get('B', set()))} "
              f"C={len(firms_by_group.get('C', set()))} "
              f"D={len(firms_by_group.get('D', set()))} "
              f"U={len(firms_by_group.get('U', set()))} "
              f"X={len(firms_by_group.get('X', set()))} firms")

    wb = Workbook()
    all_header = ["Year", "Snapshot date"] + rc_header
    write_sheet(wb.active, all_header, all_rows,
                [8, 14, 30, 44, 10, 30, 30, 26, 20, 10, 10, 8, 14, 14, 12, 60])
    wb.active.title = "All years"
    summary_header = ["Year", "Snapshot timestamp", "Table schema (columns)",
                      "Snapshot truncated at IA?", "RC rows (floor only if truncated)",
                      "Group A firms", "Group B firms", "Group C firms", "Group D firms",
                      "Group U firms", "Group X firms"]
    write_sheet(wb.create_sheet("By year summary"), summary_header, summary_rows,
                [8, 18, 18, 18, 24, 14, 14, 14, 14, 14, 14])
    out = OUT / "rate_contracts" / "RMSC_history.xlsx"
    wb.save(out)
    print(f"\nSaved: {out}")
    print(f"Years covered: {len(by_year_entries)} of {len(chosen)} picked "
          f"({len(chosen) - len(by_year_entries)} failed to download)")
    print(f"Total historical RC rows: {len(all_rows)}")


if __name__ == "__main__":
    main()
