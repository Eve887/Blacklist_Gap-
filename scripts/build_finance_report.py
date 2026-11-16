"""
build_finance_report.py - Step 5 (finance), READ-ONLY REPORT: for Nestor
Pharmaceuticals and Zee Laboratories' RMSC rate-contract lines, an
ESTIMATED 2-year contract value = rate x Annexure VIII's own "Estimated
Bid Qty. for 2years" figure for that drug code.

This is explicitly NOT a computation of money actually paid - RMSC's
own Annexure VIII quantity is itself an ESTIMATE made at tender time,
not a record of real purchase orders or deliveries. Every value in the
output is labelled "Estimated contract value (2-year), not amount
paid" (never abbreviated elsewhere in this file) per the user's
instruction.

Per the user's instruction, only NATIVE-TEXT or user-CHECKED values are
used - nothing from the Hindi OCR work (Entry 13/14) or from an
unchecked OCR deadline feeds into this. Concretely that means:
  - Rate: always native text - it comes straight from RMSC's live HTML
    rate-contract page (15_rate_contracts.py), never OCR'd.
  - 2-year estimated quantity: ONLY available for RC lines under
    NIB-14/2024/69 and NIB-10/2024 - the two tenders whose Annexure VIII
    table was read as NATIVE PDF text (pdfplumber's own table-grid
    detection, not OCR - confirmed clean on inspection, see
    PROGRESS_LOG.md). Every other tender a Nestor/Zee RC line cites
    (NIB-08/2024, NIB-01/2025, NIB-13/2024, NIB-05/2025, NIB-07/2025)
    is explicitly marked "Not available" - NIB-08/2024's tender notice
    was never found at all, NIB-01/2025 is OCR-only (not yet checked by
    the user), and NIB-13/NIB-05/NIB-07/2025 were never read in any
    form by this project - none of those qualify as "native-text or
    me-checked".

Annexure VIII page ranges (confirmed via pdfplumber's table extraction,
not assumed from page numbers alone - two tables elsewhere in each PDF
also contain the word "Code" and would have been false positives):
  NIB-14/2024/69: pages 59-95 (37 pages)
  NIB-10/2024:    pages 60-71 (12 pages)

SHARED-SUPPLIER UPPER-ESTIMATE FLAG (added after the user asked whether
Annexure VIII's quantity is per-supplier or per-tender): checked RMSC.xlsx
directly - many drugs have 2-3 suppliers under the SAME tender (L1/L2/L3
multi-vendor rate contracts, confirmed by inspection, not assumed). The
Annexure VIII quantity is the TENDER's total estimate across every
supplier of that drug, not a per-supplier figure - so assigning the
WHOLE quantity to one supplier, as this script does, is a genuine UPPER
ESTIMATE whenever a drug is shared. Every row where that applies says so
explicitly, and the value column is named accordingly (see below).

DISQUALIFICATION SCOPE (firm-wide vs product-only) - added after the
user asked for the clause's exact wording on this. RMSC's own page 9-10
clause and Annexure VII affidavit do not speak with one voice:
  - Clause I and Clause III (page 9-10) explicitly scope to the
    PRODUCT: "Bid should not be submitted for the product/products for
    which..." (I) and "...such concern/company/firm shall not be
    eligible to participate in Bid FOR SUCH PRODUCT/PRODUCTS" (III).
  - Clause II (same page) and Annexure VII item 4 (page 55-56) use
    FIRM-level wording for the identical "other state NSQ/conviction"
    trigger: "...shall ALSO not be eligible to participate IN THE BID"
    (II - no product qualifier) and "concern/company/firm does not
    stand blacklisted..." (the affidavit, same unqualified phrasing).
  This script's own read (not a legal conclusion - flagged as a
  judgment call): Clause III is the most detailed, purpose-built
  disqualification clause in the document and explicitly narrows to
  "such product/products" - Clause II's broader "firm...participate in
  the Bid" phrasing looks like loose repetition of Clause I rather than
  a deliberately wider rule. PRODUCT-ONLY is the better-supported
  reading, but the document's own drafting is genuinely ambiguous -
  both scope columns are shown so the user can judge either way.

IMPORTANT UNIT FIX, found and corrected before any number was reported:
a first version multiplied the live RC rate directly by the Annexure
VIII quantity and got figures 80-90x too large. The cause: for most
drugs, the "Packing unit" column explicitly says "Rate should be
quoted for 100 Tab" (or "10 ampules", "Single Unit", etc.) - the RATE
is PER PACK of that size, but the Annexure VIII "Estimated Bid Qty."
figure is counted in INDIVIDUAL units (tablets/ampoules/pieces), not
packs. `parse_pack_size()` reads that instruction and divides quantity
by the pack size before multiplying by rate. Verified against the
tender document's OWN "Estimated Bid Value" column (its working
estimate at tender time, a different number from this script's
output but the same order of magnitude) for 3 drugs before trusting
this for the full run - all 3 landed within ~10-18% of the tender's
own figure, which is the expected gap between a 2024/2025 planning
estimate and the rate actually contracted later. Drugs whose packing-
unit text could not be parsed (about 7% of the ~380 Annexure VIII
rows across both tenders - some genuinely ambiguous wording, some
truncated across a page boundary in the table extraction) are left
"Not available" rather than guessed.

Writes:
  output/finance/estimated_contract_value.xlsx
"""

import re

import pdfplumber
from openpyxl import Workbook, load_workbook

from common import OUT, RAW, write_sheet

FINANCE_DIR = OUT / "finance"
FINANCE_DIR.mkdir(parents=True, exist_ok=True)

ANNEXURE_VIII_PAGES = {
    "NIB14-2024": (RAW / "RMSC_Tenders" / "NIB14_2024_69_Tender14.pdf", range(59, 96)),
    "NIB10-2024": (RAW / "RMSC_Tenders" / "NIB10_2024_Tender_NIb10_DM.pdf", range(60, 72)),
}

TENDER_CODE_RE = re.compile(r"(NI[BT]-?\d+)", re.I)
TENDER_YEAR_RE = re.compile(r"\b(\d{4})(?:-\d{2,4})?\b")


def tender_short_code(tender_no):
    tender_no = tender_no or ""
    m = TENDER_CODE_RE.search(tender_no)
    y = TENDER_YEAR_RE.search(tender_no)
    if not m or not y:
        return None
    return re.sub(r"[^A-Za-z0-9]", "", m.group(1)).upper() + "-" + y.group(1)


def parse_pack_size(packing_unit_text):
    """Reads the Annexure VIII "Packing unit" column's own instruction
    ("Rate should be quoted for 100 Tab", "...for 10 ampules", "...for
    Single Unit", etc.) and returns the pack size the RATE is quoted
    per - or None if the instruction isn't confidently parsed (never
    guessed - see the module docstring's unit-fix note)."""
    t = re.sub(r"\s+", " ", packing_unit_text or "")
    m = re.search(r"quoted\s+(?:for\s+)?(\d+)\s*(amp(?:oule|ule)?s?|tab(?:let)?s?|units?|"
                 r"pieces?|cap(?:sule)?s?)\b", t, re.I)
    if m:
        return int(m.group(1))
    if re.search(r"quoted\s+(?:for\s+)?single\s+(unit|vial|pack|cartoon|carton|amp)", t, re.I):
        return 1
    return None


def load_annexure_viii_quantities():
    """Returns {tender_short_code: {drug_code: (qty_per_pack, pack_size, page_no)}} -
    qty_per_pack is Annexure VIII's own "Estimated Bid Qty. for 2years"
    figure (individual units); pack_size is what the RATE is quoted
    per (see parse_pack_size) - None if that instruction wasn't parsed,
    in which case no value will be computed for that drug (not
    guessed). Built from NATIVE PDF table extraction only, confirmed by
    spot-checking against the tender's own "Estimated Bid Value" column
    for 3 drugs before trusting the rest (see PROGRESS_LOG.md)."""
    result = {}
    for code, (pdf_path, pages) in ANNEXURE_VIII_PAGES.items():
        lookup = {}
        with pdfplumber.open(pdf_path) as pdf:
            for page_no in pages:
                page = pdf.pages[page_no - 1]
                tables = page.extract_tables()
                if not tables:
                    continue
                for row in tables[0]:
                    if not row or len(row) < 6:
                        continue
                    drug_code, qty_raw, packing_text = row[1], row[5], row[3]
                    if not drug_code or not drug_code.strip().isdigit():
                        continue   # header/continuation row, not a data row
                    qty_clean = re.sub(r"[^\d]", "", qty_raw or "")
                    if not qty_clean:
                        continue
                    pack_size = parse_pack_size(packing_text)
                    lookup[drug_code.strip()] = (int(qty_clean), pack_size, page_no)
        result[code] = lookup
        parsed = sum(1 for v in lookup.values() if v[1] is not None)
        print(f"  {code}: {len(lookup)} drug codes read from Annexure VIII "
              f"(pages {pages.start}-{pages.stop - 1}, native PDF text), "
              f"pack size parsed for {parsed}")
    return result


def parse_rate(rate_raw):
    try:
        return float(re.sub(r"[^\d.]", "", str(rate_raw)))
    except ValueError:
        return None


def load_suppliers_by_drug_tender(rc_data, rc_idx):
    """{(tender_no, drug_code): set of standardized supplier names} -
    used to flag drugs shared across more than one supplier under the
    same tender, since Annexure VIII's quantity is the TENDER's total,
    not a per-supplier figure (see module docstring)."""
    lookup = {}
    for r in rc_data:
        key = (r[rc_idx["Tender No"]], str(r[rc_idx["Drug Code"]]).strip())
        lookup.setdefault(key, set()).add(r[rc_idx["Supplier (standard)"]])
    return lookup


# Group A/B pairs from rc_vs_bans.xlsx "Strength" sheet - hardcoded here
# (only 3 rows total, all read directly from that file's own output,
# not re-derived) so the disqualification-scope scenarios below can be
# computed without re-running the classification logic.
GROUP_AB_CASES = [
    {"firm": "NESTOR PHARMACEUTICALS LTD", "tender_code": "NIB14-2024",
     "tender_no": "F.02(420)/RMSCL/PROCUREMENT/DRUG/NIB-14/2024/69", "banned_drug_code": "391"},
    {"firm": "NESTOR PHARMACEUTICALS LTD", "tender_code": "NIB08-2024",
     "tender_no": "F.02(406)/RMSCL/PROCUREMENT/DRUG/NIB-08/2024/786", "banned_drug_code": "390"},
    {"firm": "ZEE LABORATORIES LTD", "tender_code": "NIB01-2025",
     "tender_no": "F.02(422)/RMSCL/PROCUREMENT/DRUG/NIB-01/2025/871", "banned_drug_code": "NRD-486"},
]


def main():
    print("--- Reading Annexure VIII quantities (native PDF text) ---")
    annexure = load_annexure_viii_quantities()

    rc_wb = load_workbook(OUT / "rate_contracts" / "RMSC.xlsx", read_only=True)
    rc_ws = rc_wb["Rate contracts"]
    rc_rows = list(rc_ws.iter_rows(values_only=True))
    rc_header, rc_data = rc_rows[0], rc_rows[1:]
    rc_idx = {h: i for i, h in enumerate(rc_header)}

    suppliers_by_drug_tender = load_suppliers_by_drug_tender(rc_data, rc_idx)

    TARGET_FIRMS = ["NESTOR PHARMACEUTICALS LTD", "ZEE LABORATORIES LTD"]

    out_rows = []
    for r in rc_data:
        supplier = r[rc_idx["Supplier (standard)"]]
        if supplier not in TARGET_FIRMS:
            continue
        tender_no = r[rc_idx["Tender No"]]
        drug_code = str(r[rc_idx["Drug Code"]]).strip()
        rate_raw = r[rc_idx["Rate"]]
        rate = parse_rate(rate_raw)
        code = tender_short_code(tender_no)

        qty, pack_size, qty_page, estimated_value, note = None, None, None, None, ""
        if code not in annexure:
            note = f"Not available - tender {code or tender_no} was not read as native text " \
                   f"(either not found, OCR-only and unchecked, or never examined)"
        elif drug_code not in annexure[code]:
            note = f"Not available - drug code {drug_code} not found in {code}'s Annexure VIII"
        elif rate is None:
            note = f"Not available - rate '{rate_raw}' could not be parsed as a number"
        else:
            qty, pack_size, qty_page = annexure[code][drug_code]
            if pack_size is None:
                note = (f"Not available - Annexure VIII's packing-unit instruction for drug "
                        f"code {drug_code} could not be parsed, so the pack size the rate is "
                        f"quoted per is unknown (not guessed)")
            else:
                estimated_value = round(rate * (qty / pack_size), 2)
                n_suppliers = len(suppliers_by_drug_tender.get((tender_no, drug_code), {supplier}))
                if n_suppliers > 1:
                    note = (f"Shared with {n_suppliers} suppliers - value is an upper estimate "
                            f"(Annexure VIII's quantity is the TENDER's total for this drug, "
                            f"not per-supplier)")
                else:
                    note = "Estimated contract value (2-year, upper estimate), not amount paid " \
                           "- sole supplier for this drug under this tender"

        out_rows.append([
            r[rc_idx["Entry ID"]], supplier, r[rc_idx["Drug Name"]], drug_code, tender_no,
            rate_raw, qty, pack_size, estimated_value, note,
            f"Annexure VIII page {qty_page}" if qty_page else "N/A",
        ])

    header = ["Entry ID", "Supplier (standard)", "Drug Name", "Drug Code", "Tender No",
             "Rate (native text, as in RMSC.xlsx)",
             "Estimated Bid Qty. for 2 years, individual units (Annexure VIII)",
             "Pack size the Rate is quoted per (Annexure VIII's own instruction)",
             "Estimated contract value (2-year, upper estimate), not amount paid", "Note",
             "Quantity source"]

    wb = Workbook()
    write_sheet(wb.active, header, out_rows,
               [30, 24, 50, 10, 44, 14, 20, 20, 24, 70, 20])
    wb.active.title = "Estimated contract value"

    # Disqualification-scope scenarios (Group A/B cases only - see
    # GROUP_AB_CASES above and the module docstring's clause quotes).
    scenario_rows = []
    for case in GROUP_AB_CASES:
        firm_tender_lines = [r for r in out_rows
                             if r[1] == case["firm"] and r[4] == case["tender_no"]]
        whole_firm_computed = [r for r in firm_tender_lines if r[8] is not None]
        whole_firm_sum = sum(r[8] for r in whole_firm_computed)

        product_lines = [r for r in firm_tender_lines if r[3] == case["banned_drug_code"]]
        product_value = product_lines[0][8] if product_lines and product_lines[0][8] is not None else None

        scenario_rows.append([
            case["firm"], case["tender_no"], case["banned_drug_code"],
            len(firm_tender_lines), len(whole_firm_computed),
            round(whole_firm_sum, 2) if whole_firm_computed else None,
            product_value,
            "Product-only reading is the better-supported one per this script's own clause "
            "analysis (see module docstring) - shown as the more defensible estimate, but "
            "whole-firm is shown too since the document's wording is genuinely ambiguous.",
        ])

    scenario_header = ["Firm", "Tender", "Banned drug code (the Group B match)",
                      "RC lines in this tender (all)", "...with a computed value",
                      "If WHOLE-FIRM disqualification: sum of all computed RC lines in tender",
                      "If PRODUCT-ONLY disqualification: just the banned drug's line",
                      "Note"]
    write_sheet(wb.create_sheet("Disqualification scope"), scenario_header,
               scenario_rows, [26, 44, 20, 14, 14, 20, 20, 80])

    out = FINANCE_DIR / "estimated_contract_value.xlsx"
    wb.save(out)

    computed = sum(1 for r in out_rows if r[8] is not None)
    shared = sum(1 for r in out_rows if r[8] is not None and r[9].startswith("Shared with"))
    print(f"\nSaved: {out}")
    print(f"  Rows (Nestor + Zee RC lines): {len(out_rows)}")
    print(f"  Estimated contract value computed: {computed} ({shared} flagged as shared/upper estimate)")
    print(f"  Not available: {len(out_rows) - computed}")
    for firm in TARGET_FIRMS:
        firm_rows = [r for r in out_rows if r[1] == firm]
        firm_computed = [r for r in firm_rows if r[8] is not None]
        total = sum(r[8] for r in firm_computed)
        print(f"  {firm}: {len(firm_rows)} RC lines, {len(firm_computed)} computed, "
              f"sum of computed 2-year estimates = Rs {total:,.2f} (NOT a total firm value - "
              f"only the lines with available data)")
    print()
    for row in scenario_rows:
        whole_firm_str = f"Rs {row[5]:,.2f}" if row[5] is not None else "N/A"
        product_only_str = f"Rs {row[6]:,.2f}" if row[6] is not None else "N/A"
        print(f"  {row[0]} / ...{row[1][-20:]}: whole-firm={whole_firm_str}, "
              f"product-only={product_only_str}")


if __name__ == "__main__":
    main()
