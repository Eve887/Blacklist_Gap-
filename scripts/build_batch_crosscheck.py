"""
build_batch_crosscheck.py - READ-ONLY report: search the user's Project 1
CDSCO NSQ data (E:\\cdsco_mqri) for the batch numbers found in the 7
RMSC order letters (hand-read from Hindi OCR - see PROGRESS_LOG.md
Entries 13-14), to see whether the SAME batch shows up in CDSCO's own
national NSQ records.

Project 1's data IS on this computer (E:\\cdsco_mqri) - confirmed by
listing the folder, not assumed. Three already-structured/parsed
sources were searched (no new OCR or scraping done here):
  - cdsco_nsq_portal/nsq_drugs.csv      (6,956 rows, 2019-2026)
  - cdsco_nsq_portal/spurious_drugs.csv (49 rows)
  - final_package/CDSCO_NSQ_Project_Workbook.xlsx, sheet "master_alerts"
    (6,842 rows - a cleaned/deduplicated version of the two CSVs above,
    searched too in case its normalization catches something the raw
    CSVs don't)

NOT searched: the raw downloaded Alert PDFs under
cdsco_alerts/files/{central_nsq,state_nsq,combined,revised,spurious}/ -
those are PDF files, not yet turned into searchable text/rows by
Project 1 itself (its own structured output is the 3 sources above).
If the user wants those checked too, they would need to be opened/OCR'd
individually - not attempted here, logged as an open item instead of
silently skipped.

Batch numbers searched (from pdfs_to_read.xlsx's hand-read Hindi OCR
readings - every one of these is itself still "OCR - check" per rule
F, so a "not found" result here does not clear the OCR uncertainty,
and a "yes" match does not either - it only tells us whether CDSCO's
OWN database has a matching batch record):
  Zee: 722-285 (and the OCR-ambiguous variant i722.285)
  ANG: 472006, 472007
  Vee Excel: C260005, C260006
  Medipol Pharmaceutica: TMVN-143/154/157/158/161/162/163/168/169/170/171/172/173
  Medipol Pharmaceutical India Pvt Ltd: TACR-010
  Unicure: URDT0005, URDT0327, URDT0328 (all 3 OCR-ambiguous candidates)

Writes:
  output/batch_crosscheck.xlsx
"""

import csv
from pathlib import Path

from openpyxl import Workbook, load_workbook

from common import OUT, write_sheet

CDSCO_ROOT = Path("E:/cdsco_mqri")
NSQ_CSV = CDSCO_ROOT / "cdsco_nsq_portal" / "nsq_drugs.csv"
SPURIOUS_CSV = CDSCO_ROOT / "cdsco_nsq_portal" / "spurious_drugs.csv"
WORKBOOK = CDSCO_ROOT / "final_package" / "CDSCO_NSQ_Project_Workbook.xlsx"

# (batch string to search, firm, drug - as hand-read in pdfs_to_read.xlsx)
TARGETS = [
    ("722-285", "ZEE LABORATORIES LTD", "Itraconazole Cap 100mg[118]"),
    ("472006", "ANG LIFE SCIENCES INDIA LTD", "Dobutamine Inj IP 50mg/ml/250mg (vial)[492]"),
    ("472007", "ANG LIFE SCIENCES INDIA LTD", "Dobutamine Inj IP 50mg/ml/250mg (vial)[492]"),
    ("C260005", "VEE EXCEL DRUGS AND PHARMACEUTICALS PVT LTD", "Cefpodoxime Proxetil Oral Suspension [NRD-485]"),
    ("C260006", "VEE EXCEL DRUGS AND PHARMACEUTICALS PVT LTD", "Cefpodoxime Proxetil Oral Suspension [NRD-485]"),
    ("TMVN-143", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-154", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-157", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-158", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-161", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-162", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-163", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-168", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-169", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-170", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-171", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-172", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TMVN-173", "Medipol Pharmaceutica", "Multivitamin Tablets NFI Formula [394]"),
    ("TACR-010", "Medipol Pharmaceutical India Pvt Ltd", "Acyclovir Tab IP 200mg[63]"),
    ("URDT0005", "UNICURE INDIA LTD", "Glipizide And Metformin Hydrochloride Tablets USP [452]"),
    ("URDT0327", "UNICURE INDIA LTD", "Glipizide And Metformin Hydrochloride Tablets USP [452]"),
    ("URDT0328", "UNICURE INDIA LTD", "Glipizide And Metformin Hydrochloride Tablets USP [452]"),
]


def norm(s):
    return (s or "").upper().replace("-", "").replace(" ", "").replace("\n", "")


def load_csv_rows(path, batch_field):
    if not path.exists():
        return []
    with open(path, encoding="utf-8-sig") as f:
        return [(row.get(batch_field, ""), row) for row in csv.DictReader(f)]


def search_source(rows, target_batch):
    target_n = norm(target_batch)
    return [row for batch, row in rows if target_n and target_n in norm(batch)]


def main():
    if not CDSCO_ROOT.exists():
        print(f"CDSCO Project 1 data NOT found at {CDSCO_ROOT} - stopping, nothing searched.")
        return

    print(f"--- Project 1 CDSCO data found at {CDSCO_ROOT} ---")
    nsq_rows = load_csv_rows(NSQ_CSV, "str_batch_no")
    spurious_rows = load_csv_rows(SPURIOUS_CSV, "str_batch_no")
    print(f"  nsq_drugs.csv: {len(nsq_rows)} rows")
    print(f"  spurious_drugs.csv: {len(spurious_rows)} rows")

    master_rows = []
    if WORKBOOK.exists():
        wb = load_workbook(WORKBOOK, read_only=True)
        ws = wb["master_alerts"]
        rows = list(ws.iter_rows(values_only=True))
        header, data = rows[0], rows[1:]
        idx = header.index("str_batch_no")
        master_rows = [(r[idx], dict(zip(header, r))) for r in data]
        print(f"  master_alerts sheet: {len(master_rows)} rows")

    out_rows = []
    for batch, firm, drug in TARGETS:
        hits = (search_source(nsq_rows, batch) + search_source(spurious_rows, batch)
               + search_source(master_rows, batch))
        if hits:
            h = hits[0]
            state = h.get("str_reported_by_lab_or_state") or h.get("manufacturer_name", "")
            month = h.get("dt_reporting_month_year") or str(h.get("year", ""))
            match = "yes"
            note = (f"Found in CDSCO NSQ data - manufacturer listed as "
                    f"\"{h.get('str_manufactured_by') or h.get('manufacturer_name', '')}\". "
                    f"NOTE: if the reporting lab is a Rajasthan state lab, this may be the SAME "
                    f"underlying RMSC finding reported upward to CDSCO, not independent "
                    f"corroboration from elsewhere - check before treating as independent.")
        else:
            state, month, match, note = "", "", "not found", \
                "Batch string not found in any of the 3 searched CDSCO sources."
        out_rows.append([batch, firm, state, month, drug, match, note])

    header_out = ["Batch", "Firm", "State / reporting lab (if matched)",
                 "CDSCO alert month (if matched)", "Drug", "Match? (yes/no/not found)", "Note"]
    wb_out = Workbook()
    write_sheet(wb_out.active, header_out, out_rows, [14, 36, 34, 18, 44, 18, 80])
    wb_out.active.title = "Batch cross-check"
    out = OUT / "batch_crosscheck.xlsx"
    wb_out.save(out)

    found = sum(1 for r in out_rows if r[5] == "yes")
    print(f"\nSaved: {out}")
    print(f"  Batches checked: {len(out_rows)}")
    print(f"  Found in CDSCO data: {found}")
    print(f"  Not found: {len(out_rows) - found}")
    print(f"\n  NOT searched: the raw Alert PDFs under cdsco_alerts/files/ (not yet parsed "
          f"into searchable text by Project 1 itself) - open item, not silently skipped.")


if __name__ == "__main__":
    main()
