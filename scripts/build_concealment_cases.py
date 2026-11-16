"""
build_concealment_cases.py - READ-ONLY report: every CONFIRMED case in
this project's own data where a buyer's own order text says the firm
concealed/suppressed a blacklisting by ANOTHER buyer (as opposed to a
quality test failure of the CURRENT buyer's own finding).

Not part of the numbered pipeline - run it by hand whenever
master_list.xlsx changes. Creates NO new data of its own - every ban row
fact already exists in master_list.xlsx (cross-referenced by Entry ID
below); the one exception is the PMBI/Medipol row, which is NOT a normal
master_list.xlsx row at all (see CONCEALMENT_CASES below) - that fact
lives only in output/source_issues.csv and PMBI.xlsx's "Unmatched
revocations" sheet, reproduced here by hand, same as this project's other
hand-curated tables (RMSC_HINDI_READINGS, LAB_OVERRIDES).

Added 2026-11-01 per the user's own document check
(document_checks_2026-11-01.xlsx, "Concealment cases" sheet) -
5 officially-documented cases: Gujarat (GMSCL) caught 3 firms
(Unicure, Divine, Shiv Paramount) in their own debarment orders;
PMBI's revocation of Medipol's earlier action names a GMSD
concealment; Madhya Pradesh (MPPHSCL) debarred Radiant Parenterals for
hiding a Tamil Nadu (TNMSC) quality blacklisting.

Writes:
  output/concealment_cases.xlsx
"""

from openpyxl import Workbook, load_workbook

from common import OUT, write_sheet

# Each case: standardized firm name (for cross-referencing master_list.xlsx
# by Entry ID, where the row exists there), caught_by buyer, what was
# concealed (as written in the order / as found in source_issues.csv),
# and the Entry ID (or, for the one exception, a pointer to where the
# fact actually lives).
CONCEALMENT_CASES = [
    {
        "firm": "Unicure India Ltd", "caught_by": "GMSCL (Gujarat)",
        "entry_id": "GMSCL-DEBAR-15698-725-2025-26",
        "what_concealed": "\"suppression/non disclosure of debarment\"",
        "period": "26-06-2025 to 25-06-2027",
    },
    {
        "firm": "Divine Laboratories Pvt Ltd", "caught_by": "GMSCL (Gujarat)",
        "entry_id": "GMSCL-DEBAR-2025-26-2849-76",
        "what_concealed": "\"suppression/non disclosure of debarment\"",
        "period": "04-02-2026 to 03-02-2027",
    },
    {
        "firm": "Shiv Paramount Organic And Fertilizers Pvt Ltd", "caught_by": "GMSCL (Gujarat)",
        "entry_id": "GMSCL-DEBAR-14429-57-2026-27",
        "what_concealed": "\"Submission of Wrong/false/misleading Self-Attested Declaration of Debarment\"",
        "period": "31-07-2026 to 30-07-2031",
    },
    {
        "firm": "Medipol Pharmaceuticals India Pvt Ltd", "caught_by": "PMBI",
        "entry_id": None,          # NOT a normal master_list.xlsx row - see note below
        "what_concealed": "\"Concealing the fact about blacklisting of the firm by GMSD\" "
                          "(a different procurement authority, not PMBI itself)",
        "period": "Revocation dated 22-07-2022 (of PMBI's own earlier action)",
        "note": "This is PMBI's OWN REVOCATION of its earlier action against Medipol - the "
               "original PMBI entry could not be matched to any current/archived PMBI page "
               "(earliest Wayback capture of any PMBI blacklist page is 2025-09-09, 3+ years "
               "after this revocation) - see output/source_issues.csv (31-10-2026, PMBI, "
               "RevocationOrder.aspx) and output/buyers/PMBI.xlsx \"Unmatched revocations\" "
               "sheet, where the fact is actually kept. Reproduced here by reference, not as "
               "a master_list.xlsx row.",
    },
    {
        "firm": "Radiant Parenterals Ltd", "caught_by": "MPPHSCL (Madhya Pradesh)",
        "entry_id": "MPPHSCL-FIRM-273-2025",
        "what_concealed": "Did not declare that TNMSC (Tamil Nadu) had blacklisted its "
                          "Phenytoin Sodium Injection on quality grounds; false affidavit",
        "period": "27-Jan-2025 to 26-Jan-2030",
    },
]

HEADER = ["Firm", "Caught by (buyer)", "What was concealed (as written in the order)",
         "Debarred/period", "Entry ID (master_list.xlsx)", "Reason category", "Source link",
         "Note"]


def main():
    master_wb = load_workbook(OUT / "master_list.xlsx", read_only=True)
    master_ws = master_wb["Master list"]
    master_rows = list(master_ws.iter_rows(values_only=True))
    m_header, m_data = master_rows[0], master_rows[1:]
    m_idx = {h: i for i, h in enumerate(m_header)}
    by_id = {r[m_idx["Entry ID"]]: r for r in m_data}

    out_rows, missing = [], []
    for case in CONCEALMENT_CASES:
        if case["entry_id"] is None:
            out_rows.append([case["firm"], case["caught_by"], case["what_concealed"],
                             case["period"], "(none - see Note)", "Other - concealment",
                             "see output/source_issues.csv (31-10-2026, PMBI)", case["note"]])
            continue
        row = by_id.get(case["entry_id"])
        if row is None:
            missing.append(case["entry_id"])
            out_rows.append([case["firm"], case["caught_by"], case["what_concealed"],
                             case["period"], case["entry_id"] + " (NOT FOUND in master_list.xlsx)",
                             "", "", "WARNING - re-run the buyer script / merge, or check this ID"])
            continue
        out_rows.append([case["firm"], case["caught_by"], case["what_concealed"], case["period"],
                         case["entry_id"],
                         row[m_idx["Reason category (Quality / Non-supply / Other / To check)"]],
                         row[m_idx["Source link"]], case.get("note", "")])

    wb = Workbook()
    write_sheet(wb.active, HEADER, out_rows, [40, 16, 60, 30, 34, 18, 50, 70])
    wb.active.title = "Concealment cases"
    out = OUT / "concealment_cases.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Cases: {len(out_rows)}")
    if missing:
        print(f"  WARNING - {len(missing)} Entry ID(s) not found in master_list.xlsx: {missing}")


if __name__ == "__main__":
    main()
