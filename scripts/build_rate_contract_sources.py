"""
build_rate_contract_sources.py - Step 4b survey results: rate contract /
approved supplier / price-negotiation sources found across buyers.

Not part of the numbered pipeline - this is a one-time report from a
read-only survey (3 parallel research passes), not a crawler. Run again by
hand only if the survey needs updating.

Writes:
  output/rate_contract_sources.xlsx
"""

from openpyxl import Workbook

from common import OUT, write_sheet

TODAY = "31-10-2026"

HEADER = ["Rank", "Buyer", "Source / Link", "Format", "Approx entries", "Years covered",
          "Columns present", "robots.txt", "Date checked", "Notes"]

ROWS = [
    [1, "RMSC",
     "Approved Rate Lists - Drugs, Surgical and Sutures: "
     "https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-services-"
     "corporation-ltd-/en/Approved-Rate-Lists/DrugsRC.html",
     "HTML table, single page (no pagination needed)", "2,085 rows",
     "Current active contracts; Contract To Date values run out to 2028",
     "Drug Name | Code | Tender No | Supplier | Supplier Contact | Unit | Rate | GST Rate | "
     "Level (L1 marking) | Contract From Date | Contract To Date - every column asked for, "
     "plus GST and a contact field",
     "No robots.txt (404) - no restriction", TODAY,
     "By far the richest source found. Already a complete structured dataset - no PDF-opening "
     "needed at all. Clear first choice for 15_rate_contracts.py."],

    [2, "GMSCL",
     "Rate Contract: https://gmscl.gujarat.gov.in/rate-contract-eng.htm "
     "(also /rate-contract-medical-equip-and-surgi.htm, not inspected)",
     "HTML table (same ASP.NET GridView/postback pattern as the GMSCL blacklist page), "
     "with a per-row downloadable PDF work order", "13 rows seen on the page (pagination not "
     "checked - the blacklist page uses the same pattern and does paginate, so this likely does too)",
     "RO dates seen 2018-2019 in the visible sample; full range unconfirmed",
     "S.No | RC No | RO No | RO Date | Name of Item | Name of Company | Total Qty | "
     "Warranty Period | Download (PDF). No rate/price column in the HTML table itself - "
     "price is presumably in the linked work-order PDF, not confirmed by opening one.",
     "No robots.txt (404) - no restriction", TODAY,
     "Second choice. Has firm+item+qty+dates+RC/RO numbers in clean HTML, same proven "
     "postback-pagination approach as 05_gmscl.py; price needs a follow-up PDF check. "
     "An external live dashboard was also linked (gmscl.dcservices.in / a raw IP) - not "
     "inspected, likely needs a different access pattern than plain HTTP GET."],

    [3, "KMSCL",
     "Karunya Medicine Price Details: "
     "https://kmscl.kerala.gov.in/KMSCL/uploads/kcppricelist/kcppricelist140323.pdf",
     "PDF, real extractable text (pdfplumber-readable, not scanned)", "303 pages, grouped by "
     "'Company:' header (manufacturer) with multiple product rows each - likely thousands of "
     "product-level rows",
     "Single live snapshot (Report Date 03/12/2025) - a current price catalogue, not a dated "
     "historical record",
     "Sl# | Product | Unit | KCP Selling Price | MRP, grouped under each manufacturer's name. "
     "No tender number and no contract start/end date anywhere in this document.",
     "No robots.txt (404) - no restriction", TODAY,
     "Has firm+drug+price, which the lower-ranked sources below mostly lack, but no dates or "
     "tender number at all - a price list, not a rate-contract record. A second KMSCL source "
     "(\"Running Contract Details\", a category menu of ~35 equipment/reagent categories at "
     "kmscl.gov.in/.../contractslist/8/54) was found but is the Equipment Division, not drugs, "
     "and was not drilled into any category - unknown content."],

    [4, "TNMSC",
     "Empanelment / price-negotiation list: "
     "https://tnmsc.tn.gov.in/linkfiles/news_events/N2410449.pdf (the user's example; the "
     "N24xxxxx.pdf naming pattern suggests many more exist)",
     "PDF, real extractable text (pdfplumber-readable)", "~55-60 supplier rows",
     "Single tender: Enquiry 001/M(P)/Drugs/Empanelment/TNMSC/2024, dated 24.08.2024, "
     "negotiation slots 07-08.10.2024",
     "Sl.No | Name of the Supplier | No. of drugs in which L1 (lowest bidder) | Date/time of "
     "price-negotiation slot. No per-drug rate and no firm-level rupee value - just a "
     "drugs-won COUNT per firm, not a price list.",
     "No robots.txt (404) - no restriction", TODAY,
     "The example document is real and reachable, but there is still no working listing/index "
     "page to discover other N24xxxxx.pdf documents (homepage and /html/tenders.php both time "
     "out consistently - same issue logged for the blacklist work; two attempts made, not "
     "retried further). Without an index, this stays a single-document curiosity rather than a "
     "source a script could crawl."],

    [5, "APMSIDC",
     "Drugs Wing tender evaluation reports: apmsidc.ap.nic.in/wingshtml/drugswing.html, linking "
     "~40+ per-tender PDFs, some explicitly labelled 'Rate Contract' tenders (e.g. T No 167 "
     "'Branded Drugs for 2 years Rate Contract')",
     "PDF, one per tender, text-based", "~5-10 firms per tender PDF, ~40+ tender PDFs listed",
     "2022-23 through 2026-27",
     "S No | Firm Name | Firm Remarks (technical note) | Status (May be considered / May not "
     "be considered) only. No drug/item column, no rate/price, no contract dates, no tender "
     "number as its own field (only in the filename/title).",
     "Allow: / (no restriction)", TODAY,
     "This is bidder ELIGIBILITY SCREENING, not a priced rate contract or final L1 list - the "
     "least useful of the sources actually reached. A 'drugsreportarchieve.html' link exists "
     "but 404s; a fuller archive may exist elsewhere on the site, not found in this pass."],

    [6, "MPPHSCL",
     "TenderALL.aspx (tender listing, 1188 rows); a 'Rate Contracts' nav link exists but "
     "actually only holds corporate legal documents (Certificate of Incorporation, etc.), not "
     "rate-contract data",
     "HTML table (TenderALL.aspx)", "1188 rows (tender metadata only)",
     "Current/recent tenders only (dates seen Sep-Oct 2026) - not a historical archive",
     "Sr# | Tender No | Description | Category | Type | Start Date | Closing Date | "
     "TenderDocument | Corrigendum. No firm/rate/price columns on the page itself - those "
     "would be inside per-tender PDF attachments (e.g. 'Technical Evaluation Sheet (Firm "
     "wise)'), not structured data.",
     "No robots.txt (404) - no restriction", TODAY,
     "An external dashboard linked from the homepage ('MP Aushadhi Rate Contracts Details' at "
     "mpaushadhi.mp.gov.in) could NOT be checked - broken certificate chain "
     "(CERTIFICATE_VERIFY_FAILED: unable to get local issuer certificate). Did not bypass "
     "verification, per the project's standing rule that certificate checks always stay on. "
     "This is an open/unconfirmed gap, not a \"no data\" finding - may be worth a manual check "
     "in a browser, or a legitimate certificate fix, before concluding MPPHSCL has nothing."],

    [7, "PMBI",
     "Awarded Tender / Procurement PDFs: https://www.pmbi.co.in/Data/Awarded_Tender/"
     "Procurement Dept 2022_2025_13022026.pdf (+4 older Procurement-dept PDFs, 2018-19 "
     "through 2021-22)",
     "PDF - scanned images (confirmed: zero extractable text/characters on every page sampled)",
     "Unknown - 203 pages / 110MB in the largest file, cannot read without OCR",
     "2018-19 through 2022-2025 (5 separate files)", "Unknown - cannot read without OCR",
     "404 (no restriction)", TODAY,
     "Same OCR blocker already logged for RMSC's Medipol PDF and all of KSMSCL - this project's "
     "OCR environment (Tesseract) is still broken. ALSO checked: the Lok Sabha Unstarred "
     "Question 1988 (6 Dec 2024, 206 manufacturers) the user asked about - NOT FOUND. The "
     "general fact (206 manufacturers) is corroborated by other news coverage, but sansad.in's "
     "question-view pages return an identical ~3.5KB response for every query tried (a "
     "JavaScript-rendered page shell, not server-rendered content) - plain HTTP cannot read "
     "it, and no specific document URL was found or guessed."],
]


def main():
    wb = Workbook()
    write_sheet(wb.active, HEADER, ROWS, [6, 10, 50, 36, 20, 30, 50, 22, 12, 70])
    wb.active.title = "Rate contract sources"
    out = OUT / "rate_contract_sources.xlsx"
    wb.save(out)
    print(f"Saved: {out}")
    print(f"  Sources ranked: {len(ROWS)}")


if __name__ == "__main__":
    main()
