"""
build_headline_case_values.py - Step 3 (finance), READ-ONLY REPORT, for the
3 headline cases only (Zee Laboratories, Meril Endo Surgery, Nestor
Pharmaceuticals), restricted to the SPECIFIC tenders where RMSC's clause
(ii) (5(l)(II)/5(k)(ii), the whole-firm NSQ clause - see
PROGRESS_LOG.md section 3.8) applies to each firm, per the user's own
list:
  Zee:    NIB-10/2024, NIB-14/2024, NIB-01/2025
  Meril:  NIB-10/2026
  Nestor: NIB-14/2024

Every number in this report is labelled "Estimated contract value
(2-year), not amount paid" - RMSC's own Annexure VIII quantity is itself
an ESTIMATE made at tender time, not a record of real purchase orders or
deliveries. This is the SAME figure/methodology build_finance_report.py
already computes for Nestor/Zee across ALL their tenders - this script
narrows to just the clause-(ii) tenders above, ADDS Meril and the 2
SCANNED tenders (NIB-01/2025, NIB-10/2026), and adds a second, explicit
"Equal-share estimate" column (upper estimate / number of suppliers
holding an RC for the same drug code under the same tender) alongside
the existing upper-estimate concept, per the user's instruction.

USING ONLY THE FROZEN FOLDER: all buyer/case/RC facts (which firms,
which tenders, which drug codes, which rates, how many suppliers) come
from output/frozen_2026-11-07/, not the live output/ folder, so this
report stays consistent with that freeze even if the live pipeline is
re-run later. Source PDFs (data/raw/ - immutable raw evidence, never
re-derived from output/) are read directly, same as always.

PACK-SIZE FIX (reused from build_finance_report.py's parse_pack_size(),
not re-derived - see that module's own docstring for the 80-90x bug this
caught): the live RC rate is PER PACK of whatever size Annexure VIII's
own "Packing unit" column specifies ("Rate should be quoted for 100
Tab", "10 ampules", "Single Unit", etc.), but Annexure VIII's own
"Estimated Bid Qty. for 2years" figure is counted in INDIVIDUAL units -
Upper estimate = rate * (quantity / pack_size), never rate * quantity
directly.

NATIVE TEXT vs OCR (PROJECT_RULES.md rule F): NIB-14/2024 and NIB-10/2024 were
already read as clean NATIVE PDF text in build_finance_report.py
(pdfplumber's own table-grid detection, no OCR) - reused here via that
module's own load_annexure_viii_quantities(), not re-parsed. NIB-01/2025
and NIB-10/2026 are both FULLY SCANNED IMAGE PDFs (0 extractable
characters) - their Annexure VIII pages were OCR'd for this report
(resolution 250-300, English only; raw text saved to data/text/ocr/ per
rule F, same as every other OCR pass in this project) and the page
range for each was located this way, confirmed not guessed:
  NIB-01/2025 Annexure VIII: pages 59-82 (native heading "Annexure - VIII"
    on p.59, "Annexure -IX" on p.84, notes page 83 in between).
  NIB-10/2026 Annexure VIII ("List of Item(s) with Specifications"):
    pages 54-58 ("Annexure -VIII" heading on p.54, "ANNEXURE-IX" on p.59).

FOUND WHILE BUILDING THIS: unlike NIB-14/2024 and NIB-10/2024 (clean
pdfplumber table-grid extraction), these 2 scanned tables' OCR text does
NOT preserve column alignment - the SAME problem already documented for
KSMSCL's 3 scanned PDFs (PROGRESS_LOG.md, 31-10-2026 source_issues.csv
entry: "plain-text OCR does NOT preserve column alignment on dense
multi-column tables"). Spot-checking several rows found drug-code-to-
quantity pairings that do not reconcile with the RC rate at all (e.g. one
row's apparent quantity/cost pair implies a per-unit value ~3.5x the RC
rate - native-text tenders calibrated to within 10-18%, see
build_finance_report.py) - strong evidence of OCR row misalignment, not
a real figure. Per PROJECT_RULES.md rule F and the data-quality rule "preserve a
field as unknown rather than inventing a value", NO Upper/Equal-share
estimate is computed for these 2 tenders' RC lines - Quantity/Pack size/
both estimates are explicitly "Not available (OCR table alignment not
reliably parseable)", Source = "OCR - check", and the best-effort PAGE
located for that drug/item code (by keyword search against the OCR text,
hand-recorded in SCANNED_ITEM_PAGES below) is still shown, so the user
knows exactly where to look. A 10-line raw-OCR-excerpt sample (with page
numbers) is in its own sheet, per the user's instruction, for checking.

Writes:
  output/finance/headline_case_values.xlsx
"""

import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook, load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import PROJECT, RAW, write_sheet
import build_finance_report as bfr

FROZEN = PROJECT / "output" / "frozen_2026-11-07"
FINANCE_DIR = PROJECT / "output" / "finance"
FINANCE_DIR.mkdir(parents=True, exist_ok=True)

# The user's own read-by-eye check of exactly the Annexure VIII rows this
# script could not confidently OCR (see SCANNED_ITEM_PAGES above and
# checklist item 17, now CLOSED 2026-11-07) - read directly from the page
# images, not OCR, so no PROJECT_RULES.md rule F hedging is needed for these
# values. Loaded from the review file, not hand-copied into this script,
# so the "Read by eye" sheet and the Totals it feeds stay reproducible
# straight from that file.
READ_BY_EYE_FILE = PROJECT / "output" / "review" / "Finance_scanned_tenders_2026-11-07.xlsx"

# The review file writes firms as the user typed them ("Zee Laboratories",
# "Meril Endo Surgery") - common.standard_name() would NOT reliably
# reproduce CASE_TENDERS' own standardized keys ("...LTD", "...PVT LTD"),
# since the review file's text simply omits the legal suffix. An
# explicit map (only 2 firms) is safer than a fuzzy match.
READ_BY_EYE_FIRM_MAP = {
    "Zee Laboratories": "ZEE LABORATORIES LTD",
    "Meril Endo Surgery": "MERIL ENDO SURGERY PVT LTD",
}

VALUE_LABEL = "Estimated contract value (2-year), not amount paid"

# The clause-(ii) tenders this report covers, per firm - the user's own
# list (checklist/task instruction, 2026-11-07), NOT every tender the
# firm holds an RC under (e.g. Zee also holds single lines under
# NIB-13/2024, NIB-05/2025, NIB-07/2025; Nestor also holds 4 lines under
# NIB-08/2024 - all excluded, since clause (ii)'s active-on-deadline
# finding for those is not established - see rc_vs_bans.xlsx).
CASE_TENDERS = {
    "ZEE LABORATORIES LTD": ["NIB10-2024", "NIB14-2024", "NIB01-2025"],
    "MERIL ENDO SURGERY PVT LTD": ["NIB10-2026"],
    "NESTOR PHARMACEUTICALS LTD": ["NIB14-2024"],
}

SCANNED_TENDERS = {"NIB01-2025", "NIB10-2026"}

# NIB-10/2026 (sutures) Annexure VIII ("List of Item(s) with Specifications").
NIB10_2026_PDF = RAW / "RMSC_Tenders" / "NIB10_2026_Sutures_Finalbid10UBN.pdf"
NIB10_2026_PAGES = range(54, 59)

# NIB-01/2025 Annexure VIII ("List of Drugs with Specifications").
NIB01_2025_PDF = RAW / "RMSC_Tenders" / "NIB01_2025_NIB01.pdf"
NIB01_2025_PAGES = range(59, 83)

# Best-effort page location for each of Meril's (NIB-10/2026) and Zee's
# (NIB-01/2025) drug/item codes, found by keyword search against the OCR
# text saved to data/text/ocr/ (see this module's docstring) - NOT a
# confirmed quantity, only WHERE to look. "None" means the automated
# search did not confidently locate that code on any page - still logged
# here (not silently omitted) so the gap itself is visible.
SCANNED_ITEM_PAGES = {
    "NIB10-2026": {
        "R61": 57, "R74": 58, "R23": 55,
        "R6": None, "R11": None, "R15": None, "R17": None, "R18": None,
        "R20": None, "R21": None, "R22": None, "R31": None, "R34": 56,
        "R67": None, "R75": None,
    },
    "NIB01-2025": {
        "173": 59, "293": 60, "430": 60,
        "NRD-107": 65, "NRD-108": 65, "NRD-110": 65, "NRD-27": 63,
        "NRD-37": 63, "NRD-402": 71, "NRD-486": 73, "NRD-5": None,
        "NRD-520": 73, "NRD-524": 74, "NRD-56": 64, "NRD-632": 76,
        "NRD-655": 77, "NRD-813": 81, "NRD-814": None, "NRD-48": None,
        "NRD-485": None, "NRD-576": None, "NRD-670": None, "NRD-682": None,
        "NRD-683": None, "NRD-684": None, "NRD-806": 81,
    },
}


def load_rc_rows():
    """RC rows from the FROZEN folder only - see module docstring."""
    rc_wb = load_workbook(FROZEN / "rate_contracts" / "RMSC.xlsx", read_only=True)
    rc_ws = rc_wb["Rate contracts"]
    rows = list(rc_ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    idx = {h: i for i, h in enumerate(header)}
    return data, idx


def load_read_by_eye():
    """Reads the user's own "Line by line" sheet (Finance_scanned_tenders_
    2026-11-07.xlsx, output/review/) - per-line figures for the Annexure
    VIII rows this script's own OCR pass could not confidently parse.
    Returns (line_rows, totals_by_firm) where totals_by_firm is keyed by
    CASE_TENDERS' own standardized firm names (via READ_BY_EYE_FIRM_MAP,
    since the review file writes firms as the user typed them, without
    the legal suffix) to {firm: (upper_sum, equal_share_sum)}, read
    straight from that file's own 2 TOTAL rows, not recomputed here (so a
    disagreement between the file's own stated total and a sum of its
    own lines would surface in the file itself, not be silently
    reconciled by this script)."""
    wb = load_workbook(READ_BY_EYE_FILE, read_only=True)
    ws = wb["Line by line"]
    rows = list(ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    line_rows, totals_by_firm = [], {}
    for r in data:
        if r[0] is None:
            continue                                  # blank separator row
        if r[1] and str(r[1]).endswith("TOTAL"):
            firm = READ_BY_EYE_FIRM_MAP.get(r[0], r[0])
            totals_by_firm[firm] = (r[8], r[11])       # Upper estimate (Rs), Equal-share (Rs)
            continue
        line_rows.append(list(r) + [VALUE_LABEL, str(READ_BY_EYE_FILE.relative_to(PROJECT))])
    return line_rows, totals_by_firm


def count_suppliers(rc_data, rc_idx, tender_no, drug_code):
    """Distinct suppliers holding an RC for this exact drug code under
    this exact tender - computed from EVERY row in RMSC.xlsx (not just
    the 3 headline firms), since another, unrelated supplier sharing the
    same drug code still shares Annexure VIII's single tender-wide
    quantity estimate."""
    suppliers = {r[rc_idx["Supplier (standard)"]] for r in rc_data
                 if r[rc_idx["Tender No"]] == tender_no
                 and str(r[rc_idx["Drug Code"]]).strip() == drug_code}
    return suppliers


def main():
    print("--- Loading native-text Annexure VIII (NIB-14/2024, NIB-10/2024) ---")
    native_annexure = bfr.load_annexure_viii_quantities()

    print("--- Loading RC data from the FROZEN folder ---")
    print(f"  {FROZEN}")
    rc_data, rc_idx = load_rc_rows()

    out_rows = []
    for firm, codes in CASE_TENDERS.items():
        for r in rc_data:
            if r[rc_idx["Supplier (standard)"]] != firm:
                continue
            tender_no = r[rc_idx["Tender No"]]
            code = bfr.tender_short_code(tender_no)
            if code not in codes:
                continue                              # not a clause-(ii) tender for this firm

            drug_code = str(r[rc_idx["Drug Code"]]).strip()
            rate_raw = r[rc_idx["Rate"]]
            rate = bfr.parse_rate(rate_raw)
            suppliers = count_suppliers(rc_data, rc_idx, tender_no, drug_code)
            n_suppliers = len(suppliers)

            qty, pack_size, page, source, upper, equal_share, note = (
                None, None, None, None, None, None, "")

            if code in SCANNED_TENDERS:
                source = "OCR - check"
                page = SCANNED_ITEM_PAGES.get(code, {}).get(drug_code)
                page_str = f"page {page}" if page else "not located by automated OCR search"
                note = (f"Not available from OCR alone - Annexure VIII for this tender is a "
                        f"SCANNED image PDF; OCR text does not reliably preserve column "
                        f"alignment for this dense table (PROJECT_RULES.md rule F - see module "
                        f"docstring), so no quantity was computed here. Candidate OCR location: "
                        f"{page_str} (data/text/ocr/). RESOLVED 2026-11-07 - the user read this "
                        f"line directly from the page image; see the 'Read by eye' sheet for the "
                        f"checked figure.")
            elif code in native_annexure and rate is not None:
                source = "native text"
                if drug_code in native_annexure[code]:
                    qty, pack_size, page = native_annexure[code][drug_code]
                    if pack_size is None:
                        note = (f"Not available - Annexure VIII's packing-unit instruction for "
                                f"drug code {drug_code} could not be parsed (pack size unknown, "
                                f"not guessed)")
                    else:
                        upper = round(rate * (qty / pack_size), 2)
                        equal_share = round(upper / n_suppliers, 2)
                        if n_suppliers > 1:
                            note = (f"{VALUE_LABEL}. Upper estimate assigns the WHOLE tender "
                                     f"quantity to this one supplier (an upper bound - "
                                     f"{n_suppliers} suppliers hold an RC for this drug code "
                                     f"under this tender); Equal-share estimate divides it "
                                     f"evenly across all {n_suppliers}.")
                        else:
                            note = (f"{VALUE_LABEL}. Sole supplier for this drug code under this "
                                     f"tender - Upper estimate and Equal-share estimate are the "
                                     f"same figure.")
                else:
                    note = f"Not available - drug code {drug_code} not found in {code}'s Annexure VIII"
            elif rate is None:
                source = "native text"
                note = f"Not available - rate '{rate_raw}' could not be parsed as a number"
            else:
                source = "native text"
                note = f"Not available - tender {code} Annexure VIII was not loaded"

            out_rows.append([
                r[rc_idx["Entry ID"]], firm, r[rc_idx["Drug Name"]], drug_code, tender_no,
                rate_raw, None,                  # Unit text filled in below, after the loop
                pack_size, qty, n_suppliers,
                upper, equal_share, source,
                f"page {page}" if page else "N/A", note,
            ])

    # The "Unit" column needs the actual packing-unit TEXT (not just the
    # parsed number) - filled in now, in a second pass, reading it straight
    # out of the same Annexure VIII table cell for native-text rows (not
    # re-parsed into a number, since parse_pack_size() only keeps the
    # numeric pack size, discarding the instruction text itself). Scanned-
    # tender rows get "Not available" (same reason as Quantity/Pack size).
    unit_text_by_code = _load_unit_text(native_annexure.keys())
    for row in out_rows:
        tender_no = row[4]
        drug_code = row[3]
        code = bfr.tender_short_code(tender_no)
        row[6] = unit_text_by_code.get((code, drug_code), "Not available" if code in SCANNED_TENDERS
                                        else "Not available")

    header = ["Entry ID", "Supplier (standard)", "Drug Name", "Drug Code", "Tender No",
              "Rate (native text, RMSC.xlsx)", "Unit (Annexure VIII packing-unit instruction)",
              "Pack size (parsed from Unit)",
              "Annexure VIII 2-year estimated quantity (individual units)",
              "Number of suppliers (same drug code, same tender)",
              f"Upper estimate - {VALUE_LABEL}",
              f"Equal-share estimate - {VALUE_LABEL}",
              "Source: native text / OCR - check", "Source page", "Note"]

    wb = Workbook()
    write_sheet(wb.active, header, out_rows,
                [26, 28, 50, 12, 46, 14, 46, 12, 16, 14, 20, 20, 18, 10, 90])
    wb.active.title = "RC lines"

    # Totals per firm and per tender - only over rows with a computed value
    # (never a count of "Not available" rows masquerading as zero).
    def _totals_row(firm, label, rows):
        n_computed = sum(1 for r in rows if r[10] is not None)
        if n_computed == 0:
            return [firm, label, len(rows), 0, "N/A - no RC line in this scope has a computed "
                    "value (see Note column in 'RC lines')", "N/A", VALUE_LABEL]
        u = sum(r[10] for r in rows if r[10] is not None)
        e = sum(r[11] for r in rows if r[11] is not None)
        return [firm, label, len(rows), n_computed, round(u, 2), round(e, 2), VALUE_LABEL]

    totals_rows = []
    for firm, codes in CASE_TENDERS.items():
        firm_rows = [r for r in out_rows if r[1] == firm]
        totals_rows.append(_totals_row(firm, "ALL clause-(ii) tenders (native text only)", firm_rows))
        for code in codes:
            tender_rows = [r for r in firm_rows if bfr.tender_short_code(r[4]) == code]
            if not tender_rows:
                continue
            totals_rows.append(_totals_row(firm, code, tender_rows))

    totals_header = ["Firm", "Tender (ALL = sum across its clause-(ii) tenders)",
                      "RC lines", "...with a computed value",
                      f"Sum of Upper estimates - {VALUE_LABEL}",
                      f"Sum of Equal-share estimates - {VALUE_LABEL}", "Label"]

    # Read-by-eye sheet + figures (checklist item 17, CLOSED 2026-11-07 -
    # see module docstring/READ_BY_EYE_FILE) - loaded from the user's own
    # review file, not hand-typed into this script.
    print("--- Loading the read-by-eye scanned-tender figures ---")
    print(f"  {READ_BY_EYE_FILE}")
    read_by_eye_rows, read_by_eye_totals = load_read_by_eye()

    read_by_eye_header = ["Firm", "Tender", "Code", "Item", "Rate is quoted",
                          "Rate (Rs, excl. GST)", "2-year estimated quantity (units)",
                          "Units per rate",
                          f"Upper estimate (Rs) - {VALUE_LABEL}", "Upper estimate (Rs lakh)",
                          "Suppliers sharing this code",
                          f"Equal-share estimate (Rs) - {VALUE_LABEL}",
                          "Rajasthan's own estimated value (Rs) (a check figure, not used in "
                          "the estimate)", "Annexure VIII page", "Read status", "Label", "Source"]
    write_sheet(wb.create_sheet("Read by eye"), read_by_eye_header, read_by_eye_rows,
                [20, 22, 10, 56, 18, 16, 20, 12, 20, 16, 16, 20, 20, 12, 36, 50, 60])

    # Combined totals (native text, where read natively, PLUS the read-by-
    # eye figures for the 2 scanned tenders) - the authoritative per-firm
    # figure this report is FOR, per the user's instruction to "update
    # totals". Nestor has no scanned tender in its clause-(ii) scope, so
    # its combined total equals its native-only total exactly.
    combined_rows = []
    for firm in CASE_TENDERS:
        native_upper = sum(r[10] for r in out_rows if r[1] == firm and r[10] is not None)
        native_equal = sum(r[11] for r in out_rows if r[1] == firm and r[11] is not None)
        eye_upper, eye_equal = read_by_eye_totals.get(firm, (0, 0))
        combined_rows.append([
            firm, round(native_upper, 2), round(eye_upper or 0, 2),
            round(native_upper + (eye_upper or 0), 2), round(native_equal, 2),
            round(eye_equal or 0, 2), round(native_equal + (eye_equal or 0), 2), VALUE_LABEL,
        ])
    combined_header = ["Firm", "Upper estimate - native text only", "Upper estimate - read by "
                       "eye (scanned tenders)", f"Upper estimate - COMBINED (authoritative) - "
                       f"{VALUE_LABEL}", "Equal-share estimate - native text only",
                       "Equal-share estimate - read by eye (scanned tenders)",
                       f"Equal-share estimate - COMBINED (authoritative) - {VALUE_LABEL}", "Label"]

    write_sheet(wb.create_sheet("Totals"), totals_header, totals_rows,
                [30, 46, 10, 16, 26, 26, 50])
    # The authoritative, combined (native text + read by eye) figure for
    # each firm - see this sheet, not "Totals" (native text only) or
    # "Read by eye" (scanned tenders only) alone, for the final number.
    write_sheet(wb.create_sheet("Combined totals"), combined_header, combined_rows,
                [30, 26, 30, 30, 26, 32, 32, 50])

    # 10-line sample of the scanned-tender OCR, with page numbers, per the
    # user's instruction - for checking the OCR itself, not tied only to
    # Zee/Meril's own drug codes, so the sample shows representative rows
    # across both tenders' tables.
    sample_rows = _build_ocr_sample()
    write_sheet(wb.create_sheet("Scanned tender OCR sample"),
                ["Tender", "Page", "Raw OCR excerpt (PROJECT_RULES.md rule F - OCR, not checked)"],
                sample_rows, [16, 10, 110])

    out = FINANCE_DIR / "headline_case_values.xlsx"
    wb.save(out)

    print(f"\nSaved: {out}")
    for firm in CASE_TENDERS:
        firm_rows = [r for r in out_rows if r[1] == firm]
        computed = sum(1 for r in firm_rows if r[10] is not None)
        print(f"  {firm}: {len(firm_rows)} RC lines, {computed} with a computed estimate")


def _load_unit_text(native_codes):
    """Re-reads the Annexure VIII "Packing unit" column's own TEXT (not
    the parsed number) for NIB-14/2024 and NIB-10/2024, keyed by
    (tender_short_code, drug_code) - a second, light pass over the same
    native-text pages bfr.load_annexure_viii_quantities() already read,
    so the report can show the instruction itself, not just the pack
    size parsed out of it."""
    result = {}
    for code in native_codes:
        pdf_path, pages = bfr.ANNEXURE_VIII_PAGES[code]
        with pdfplumber.open(pdf_path) as pdf:
            for page_no in pages:
                page = pdf.pages[page_no - 1]
                tables = page.extract_tables()
                if not tables:
                    continue
                for row in tables[0]:
                    if not row or len(row) < 6:
                        continue
                    drug_code, packing_text = row[1], row[3]
                    if not drug_code or not drug_code.strip().isdigit():
                        continue
                    result[(code, drug_code.strip())] = re.sub(r"\s+", " ", packing_text or "").strip()
    return result


def _build_ocr_sample():
    """10 representative lines (with page numbers) from the 2 scanned
    tenders' Annexure VIII OCR text - a few whole paragraphs' worth per
    page so a line is actually checkable against the PDF, not a single
    disconnected word."""
    picks = [
        ("NIB-10/2026", 54, "range1"), ("NIB-10/2026", 55, "range1"),
        ("NIB-10/2026", 56, "range1"), ("NIB-10/2026", 57, "range1"),
        ("NIB-10/2026", 58, "range1"),
        ("NIB-01/2025", 59, "range1"), ("NIB-01/2025", 63, "range1"),
        ("NIB-01/2025", 65, "range1"), ("NIB-01/2025", 73, "range1"),
        ("NIB-01/2025", 81, "range1"),
    ]
    rows = []
    for tender, page, _ in picks:
        fname = ("NIB10_2026_Sutures_Finalbid10UBN" if tender == "NIB-10/2026"
                 else "NIB01_2025_NIB01")
        path = PROJECT / "data" / "text" / "ocr" / f"{fname}_p{page}.txt"
        text = path.read_text(encoding="utf-8") if path.exists() else "(not found)"
        excerpt = re.sub(r"\s+", " ", text).strip()[:400]
        rows.append([tender, page, excerpt])
    return rows


if __name__ == "__main__":
    main()
