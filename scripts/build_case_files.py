"""
build_case_files.py - one case file per firm of interest: every ban
record (any buyer), every RMSC rate-contract line (if any), and any
outside evidence the user has supplied.

Not part of the numbered pipeline - run by hand whenever a case firm's
underlying data changes. Creates NO new data of its own for the
"Bans"/"RMSC RC lines" sheets - every fact there already exists in
master_list.xlsx / rate_contracts/RMSC.xlsx; this only collects and
presents it per firm. The "Outside evidence" sheet is the one place
with genuinely new content - a citation the user supplied by hand, not
found or verified by this script.

NAME VARIANTS ARE NEVER MERGED. Where RMSC's own record of a firm omits
the legal suffix other buyers include (e.g. "NESTOR PHARMACEUTICALS" vs
"NESTOR PHARMACEUTICALS LTD", "ZEE LABORATORIES" vs "ZEE LABORATORIES
LTD" - both confirmed in master_list.xlsx), both standardized names are
INCLUDED in the case file's "Bans" sheet with their own exact Entry ID
and source link, but flagged in a separate "Possibly related name
variants" sheet rather than silently treated as the same row - a human
decides whether they're the same firm, same as every other matching
question in this project.

Writes:
  output/cases/<Firm>.xlsx - one per entry in CASES below
"""

from datetime import date

from openpyxl import Workbook, load_workbook

from common import OUT, write_sheet

TODAY = date.today().strftime("%d-%m-%Y")
CASES_DIR = OUT / "cases"
CASES_DIR.mkdir(parents=True, exist_ok=True)

# Each case: the file name, the EXACT standardized name(s) to collect
# bans/RC lines for (plural to catch an unmerged variant - see module
# docstring), and any outside evidence supplied by the user (empty list
# if none).
CASES = {
    "Nestor": {
        "names": ["NESTOR PHARMACEUTICALS LTD", "NESTOR PHARMACEUTICALS"],
        "outside_evidence": [
            ["The Tribune India (+ CAG Haryana audit report)", "14-11-2024 (article); CAG report "
             "'Public Health Infrastructure and Management of Health Services', tabled in the "
             "Haryana Vidhan Sabha",
             "CAG found Haryana Medical Services Corporation Ltd (HMSCL) purchased drugs worth "
             "Rs 1.52 crore from Nestor Pharmaceuticals (May-Dec 2019) for folic acid-ferrous "
             "sulphate tablets - AFTER Gujarat Medical Services Corporation Ltd (GMSCL) had "
             "already blacklisted Nestor for 3 years in Feb 2017 for supplying the SAME drug, "
             "and AFTER HMSCL itself separately blacklisted Nestor for 3 years in Sep 2018 for "
             "concealing the GMSCL action. A confirmed, independently-audited coordination gap -"
             " the same drug (folic acid-ferrous sulphate) this project's own RMSC/Nestor Group B "
             "finding was about before it was retracted on dosage-form grounds (see Entry 16) -"
             " a different buyer pair (GMSCL/HMSCL, not GMSCL/RMSC), genuinely confirmed, not "
             "retracted.",
             "https://www.tribuneindia.com/news/haryana/haryana-bought-drugs-from-blacklisted-firm-cag/",
             "Found via web search 2026-10-31, NOT supplied by the user with a link (unlike the "
             "Jackson/Newslaundry case) - the user referred to 'the CAG Haryana finding' assuming "
             "it was already known; it was not, and was located and read for this entry. The "
             "Rs 1.52 crore figure is the CAG's/press's OWN reported figure, not computed by this "
             "project (per the standing 'no rupee values' instruction - cited, not calculated)."],
        ],
        "extra_facts": [],
    },
    "Zee": {
        "names": ["ZEE LABORATORIES LTD", "ZEE LABORATORIES"],
        "outside_evidence": [],
        "extra_facts": [],
    },
    "Jackson": {
        # "JACKSON LABORATORIES PVT LTD" added 2026-11-01 - the Karnataka full-list
        # rewrite (16_ksmscl.py, replacing the old OCR fragment "Jackson") uses the
        # firm's full legal name, which standard_name() does not collapse to the
        # shorter "JACKSON LABORATORIES" PMBI already used - see "Possibly related
        # name variants" sheet below, not silently merged.
        "names": ["JACKSON LABORATORIES", "JACKSON LABORATORIES PVT LTD"],
        "outside_evidence": [
            ["Newslaundry", "17-01-2024",
             "Reports Karnataka (KSMSCL) blacklisted Jackson Laboratories for 5 years from "
             "2023, while Delhi's Central Procurement Agency (CPA) continued procuring from "
             "it, despite multiple failed quality tests.",
             "https://www.newslaundry.com/amp/story/2024%2F01%2F17%2Fexclusive-aap-govt-"
             "procured-drugs-from-2-blacklisted-pharma-firms-with-many-failed-quality-tests",
             "Supplied by the user 2026-10-31 - NOT independently verified by this script "
             "(no automated fetch/check was done against this URL)."],
        ],
        "extra_facts": [],
    },
    "Meril": {
        "names": ["MERIL ENDO SURGERY PVT LTD", "MERIL ENDO SURGERY"],
        "outside_evidence": [],
        "extra_facts": [
            ["GMSCL ban period vs RMSC NIB-10/2026 sutures deadline",
             "GMSCL ban GMSCL-DEBAR-18621-652-2025-26: 16-07-2025 to 15-07-2028. RMSC "
             "NIB-10/2026 (sutures) bid-submission deadline: 06-05-2026, 6:00 PM - CONFIRMED by "
             "the user, by eye, against pages 1-3 of the tender PDF, 2026-11-07 (no longer "
             "OCR-unchecked - PROJECT_RULES.md rule F satisfied; see build_rc_vs_bans.py "
             "TENDER_DEADLINES). 16-07-2025 <= 06-05-2026 <= 15-07-2028: ACTIVE on the "
             "deadline - computed directly, not guessed."],
            ["Which lab found the sterility failure (GMSCL ban)",
             "Food and Drug Laboratory, Vadodara (Gujarat's own statutory drug-testing body) - "
             "read directly from the full GMSCL order PDF (Debar_Circular-23_-_Meril_Endo-"
             "_Suture-1.pdf), native text, not OCR. The live GMSCL summary page itself never "
             "names a lab - this fact only exists in the full order. Logged as a Q1-grade "
             "finding (common.py LAB_OVERRIDES in 05_gmscl.py) since this IS a named statutory "
             "lab, even though the live page's own text doesn't say so."],
            ["RMSC relevance", "Meril has 17 RMSC RC lines (sutures) but RMSC's own eligibility "
             "clause applies to DRUG tenders with the SAME wording confirmed in the Surgical & "
             "Sutures tender too (NIB-10/2026, page 9) - see PROGRESS_LOG.md. The GMSCL ban "
             "overlapping the RMSC Drug-tender Group A/B check does not apply here directly "
             "(Meril is not a Drug-tender RC supplier), but the SAME policy-gap question "
             "applies to the Surgical/Sutures side of RMSC's business, not yet built as its "
             "own check."],
        ],
    },
}

BAN_HEADER = ["Entry ID", "Firm name (as written)", "Firm name (standard)", "Buyer",
             "Buyer state", "Level (firm / product)", "Product(s)", "Blacklisted from",
             "Blacklisted to", "Period", "Reason (as written)",
             "Reason category (Quality / Non-supply / Other / To check)",
             "Quality level (Q1 = statutory lab named, Q2 = other/no lab)",
             "Lab that declared NSQ",
             "Court stay or revoked?", "Source link", "Status"]

RC_HEADER = ["Entry ID", "Drug Name", "Drug Code", "Tender No", "Supplier (standard)",
            "Rate", "GST Rate", "Level", "Contract From", "Contract To"]


def main():
    master_wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    master_ws = master_wb["Master list"]
    master_rows = list(master_ws.iter_rows(values_only=True))
    m_header, m_data = master_rows[0], master_rows[1:]
    m_idx = {h: i for i, h in enumerate(m_header)}

    rc_wb = load_workbook(OUT / "rate_contracts" / "RMSC.xlsx", read_only=True)
    rc_ws = rc_wb["Rate contracts"]
    rc_rows = list(rc_ws.iter_rows(values_only=True))
    rc_header_in, rc_data = rc_rows[0], rc_rows[1:]
    rc_idx = {h: i for i, h in enumerate(rc_header_in)}

    for case_name, case in CASES.items():
        names = case["names"]
        ban_rows = [
            [r[m_idx["Entry ID"]], r[m_idx["Firm name (as written)"]], r[m_idx["Firm name (standard)"]],
             r[m_idx["Buyer"]], r[m_idx["Buyer state"]], r[m_idx["Level (firm / product)"]],
             r[m_idx["Product(s)"]], r[m_idx["Blacklisted from"]], r[m_idx["Blacklisted to"]],
             r[m_idx["Period"]], r[m_idx["Reason (as written)"]],
             r[m_idx["Reason category (Quality / Non-supply / Other / To check)"]],
             r[m_idx["Quality level (Q1 = statutory lab named, Q2 = other/no lab)"]],
             r[m_idx["Lab that declared NSQ"]],
             r[m_idx["Court stay or revoked?"]], r[m_idx["Source link"]], r[m_idx["Status"]]]
            for r in m_data if r[m_idx["Firm name (standard)"]] in names
        ]
        rc_rows_out = [
            [r[rc_idx["Entry ID"]], r[rc_idx["Drug Name"]], r[rc_idx["Drug Code"]],
             r[rc_idx["Tender No"]], r[rc_idx["Supplier (standard)"]], r[rc_idx["Rate"]],
             r[rc_idx["GST Rate"]], r[rc_idx["Level"]], r[rc_idx["Contract From"]],
             r[rc_idx["Contract To"]]]
            for r in rc_data if r[rc_idx["Supplier (standard)"]] in names
        ]

        wb = Workbook()
        write_sheet(wb.active, BAN_HEADER, ban_rows,
                   [26, 34, 30, 10, 14, 16, 44, 13, 13, 14, 30, 14, 14, 30, 30, 50, 12])
        wb.active.title = "Bans (all buyers)"
        write_sheet(wb.create_sheet("RMSC RC lines"), RC_HEADER, rc_rows_out,
                   [34, 44, 10, 40, 24, 10, 10, 8, 14, 14])

        if len(names) > 1:
            variant_rows = [[n, sum(1 for r in ban_rows if r[2] == n),
                            "NOT merged - different standardized name from RMSC's own record "
                            "than other buyers use for what may be the same firm. A human "
                            "decision, not an automated one."] for n in names]
            write_sheet(wb.create_sheet("Possibly related name variants"),
                       ["Standardized name used", "Ban rows under this name", "Note"],
                       variant_rows, [30, 20, 80])

        if case["outside_evidence"]:
            write_sheet(wb.create_sheet("Outside evidence"),
                       ["Source", "Date", "What it reports", "Link", "Note"],
                       case["outside_evidence"], [20, 14, 70, 70, 60])

        if case.get("extra_facts"):
            write_sheet(wb.create_sheet("Case-specific findings"),
                       ["Question", "Finding"], case["extra_facts"], [40, 110])

        out = CASES_DIR / f"{case_name}.xlsx"
        wb.save(out)
        print(f"Saved: {out}")
        print(f"  Ban rows: {len(ban_rows)}  (buyers: {sorted(set(r[3] for r in ban_rows))})")
        print(f"  RMSC RC lines: {len(rc_rows_out)}")
        if len(names) > 1:
            print(f"  Name variants NOT merged: {names}")
        if case["outside_evidence"]:
            print(f"  Outside evidence rows: {len(case['outside_evidence'])}")
        if case.get("extra_facts"):
            print(f"  Case-specific findings: {len(case['extra_facts'])}")
        print()


if __name__ == "__main__":
    main()
