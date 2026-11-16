# Case file: Meril Endo Surgery Pvt Ltd

Part of "The Blacklist Gap" review. Buyers may have reasons not visible in
public records; the firm was not contacted for comment. Every finding
below is a possible breach of a tender clause as worded, not a finding of
fraud or illegal conduct.

## Headline finding

Gujarat's GMSCL debarred Meril's Non-Absorbable Black Braided (BB) Silk
surgical suture (needle 1-0) for 3 years after the Food and Drug
Laboratory, Vadodara - a statutory lab (Q1) - reported it "Not of Standard
Quality" on a sterility test. Meril separately holds 15 rate-contract
lines with Rajasthan's RMSC under the sutures tender NIB-10/2026, active
on that tender's own bid deadline. Under RMSC's clause 5(k)(ii) (the
sutures tender's own version of clause 5(l)(II), confirmed word-for-word
identical in the relevant wording), this is a **possible breach of the
tender clause as worded**.

## Timeline

| Date | Event | Source |
|---|---|---|
| 28-02-2020 | GMSCL floats tender enquiry GMSCL/D-662/RC/2019-20 (technical bid due 03-06-2020) - the tender later leading to the rate contract found NSQ | GMSCL order, page 1, table column (2) |
| 17-05-2021 | GMSCL executes the rate contract (GMSCL/Drugs/RC/662-4073/D-18/15456-57/2021-22, valid upto 30-04-2025) with Meril for the item later found NSQ. *Corrected 2026-11-07: an earlier draft of this case file read this date as "28-02-2024" and the tender-enquiry number as "GMSCUD-662/RC/2019-20" - both wrong. The "28-02-2024" figure was a transcription error; "GMSCUD" is an independently-confirmed font/character-encoding artifact in this specific PDF's own embedded text layer (both pdfplumber and PyMuPDF extract the same "GMSCUD" from the native text; the rendered page image itself clearly reads "GMSCL/D-662") - not a typo on the buyer's part, but also not what either extraction library returns.* | GMSCL order, page 1, table column (3) |
| 01-09-2023 to 31-08-2024 | MPPHSCL bans a different Meril product for Non-supply (expired, not a quality finding) | `MPPHSCL-PRODUCT-6972-2023` |
| 28-01-2025 | Food and Drug Laboratory, Vadodara reports the suture sample (test report FDL/GMSCL-PDT/4359/2025) "Not of Standard Quality" (sterility test) | GMSCL order, page 1 |
| 11-02-2025 | GMSCL serves a show-cause notice on Meril | GMSCL order, page 1 |
| 06-11-2024 | MPPHSCL orders Meril's firm-level blacklisting (order 7876/F605/2024) for failing to execute an agreement (a rate-quoting error on a single-foil vs per-pack basis, EMD Rs 2,00,000 forfeited) - **a separate action, not a quality finding** | `MPPHSCL-FIRM-7876-2024` |
| 20-01-2025 | MPPHSCL's Commissioner, on appeal (order 167), SETS ASIDE the firm-level part of the 06-11-2024 order - Meril remains blacklisted only for item MIS93 (the same polyester-braided suture item), 2 years from 06-11-2024. Confirmed (2026-11-07) that `MPPHSCL-PRODUCT-166-2024` - a separate record on MPPHSCL's own Product list - is this SAME surviving action: its own PDF is byte-for-byte identical (SHA-256 hash match) to the original order PDF. Not a quality finding either way - shown here for completeness, not part of the headline case. | `MPPHSCL-PRODUCT-166-2024`; see `output/source_issues.csv`, 07-11-2026 |
| **16-07-2025** | GMSCL debarment takes effect - 3 years, to 15-07-2028 | GMSCL order no. GMSCL/SD/Debar/23/652/2025-26, page 2 |
| **06-05-2026, 6 PM** | RMSC's Rajasthan sutures tender NIB-10/2026 bid deadline - confirmed by the researcher reading pages 1-3 of the tender PDF directly (no longer OCR-derived) - GMSCL ban active (16-07-2025 to 15-07-2028 covers this date) | `data/raw/RMSC_Tenders/NIB10_2026_Sutures_Finalbid10UBN.pdf`, pages 1-3 |
| 20-08-2026 to 27-08-2026 | Meril's 15 RMSC rate-contract lines under NIB-10/2026 begin (various start dates per line, all AFTER the bid deadline above). *Corrected 2026-11-07: an earlier draft of this case file read these dates as 2025 - the source data's own Contract From dates are 20-AUG-2026/27-AUG-2026, a year later.* | `output/frozen_2026-11-07/rate_contracts/RMSC.xlsx` |

## Every source document, with page numbers

- GMSCL debarment order (Meril, GMSCL/SD/Debar/23/652/2025-26): page 1
  (findings), page 2 (operative order and distribution list) -
  `data/raw/GMSCL/candidate_gap_pdfs/GMSCL-DEBAR-18621-652-2025-26.pdf` -
  https://gmscl.gujarat.gov.in/Portal/Firm_BlackList/Debar_Circular-23_-_Meril_Endo-_Suture-1.pdf
- MPPHSCL firm-level order (7876/2024) and its appeal (order 167,
  20-01-2025): `output/frozen_2026-11-07/buyers/MPPHSCL.xlsx`, Entry ID
  `MPPHSCL-FIRM-7876-2024` - https://mpphscl.mp.gov.in/Files/PDF/a05bc405-f12d-4b39-94bb-848c3f118abe_3_order meril.pdf
- MPPHSCL product-level record for the same surviving action: Entry ID
  `MPPHSCL-PRODUCT-166-2024` - https://mpphscl.mp.gov.in/Files/PDF/54ee812a-afa0-405f-8f1a-d49a9da5b552_1_order meril.pdf
- RMSC sutures tender NIB-10/2026: pages 1-3 (bid deadline), page 9
  (clause 5(k)), pages 54-58 (Annexure VIII, "List of Item(s) with
  Specifications") - `data/raw/RMSC_Tenders/NIB10_2026_Sutures_Finalbid10UBN.pdf`
- RMSC's own rate-contract list (15 lines under NIB-10/2026): `output/
  frozen_2026-11-07/rate_contracts/RMSC.xlsx` - https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-services-corporation-ltd-/en/Approved-Rate-Lists/DrugsRC.html
- Full line-by-line detail, read by eye (all 15 lines), and both
  estimates: `output/finance/headline_case_values.xlsx` ("Read by eye" and
  "Combined totals" sheets); the underlying hand-read figures themselves
  are in `output/review/Finance_scanned_tenders_2026-11-07.xlsx`.

## Not part of the headline finding: outside CDSCO's scope

The contagion test in `output/analysis/contagion_test.xlsx` (section 4(f)
of the main report) could not include Meril at all - its ban is for a
surgical suture, a medical device, outside the scope of CDSCO's drug-only
NSQ alert portal.

## Estimated contract value

Rs 0.99 crore (equal-share estimate) to Rs 2.28 crore (upper estimate),
2-year estimated contract value across all 15 rate-contract lines, all
read by eye from the scanned tender's own Annexure VIII page images (the
automated OCR pass could not reliably keep this dense table's columns
aligned) - **not an amount paid**. Full line-by-line breakdown:
`output/finance/headline_case_values.xlsx`.
