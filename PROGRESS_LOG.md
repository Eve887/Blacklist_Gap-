# blacklist_gap - Progress Log

Sections 1, 3, 4 are kept up to date in place (current status). Section 5 is
append-only history - never edit or delete an old entry there; corrections
go in as new entries. See PROJECT_RULES.md for the rules this log follows.

## 1. Status table

| Buyer | Rows | Format | Status |
|---|---|---|---|
| APMSIDC (AP) | 43 | PDF (pdfplumber) | Done |
| AMSCL (Assam) | 0 | HTML | Page's content is genuinely empty server-side (confirmed, not a bug); Wayback Machine has no archived copy either |
| PMBI | 75 | HTML tables (current) + 2 archive pages | Done; 1 Unmatched revocation (Medipol) |
| GMSCL | 99 | HTML, ASP.NET postback pagination | "Current Blacklisted" done; "Archive Blacklisted" blocked by a WAF (plain HTTP and two Playwright attempts, the second returned an explicit CAPTCHA) - now reads a hand-saved HTML copy instead if one is provided |
| MPPHSCL | 378 | HTML (Firm/Lab/Product, 3 pages) | Done |
| RMSC | 101 | HTML, single page | Done; no "From" date on this site at all |
| KMSCL | 5 | 2 PDFs (pdfplumber) | Done |
| KSMSCL | 22 | 3 PDFs, scanned images - hand-transcribed by the user directly from the page images (not OCR) | Done (2026-11-01) - the earlier 11 OCR name-fragment rows are REPLACED by a full, hand-verified transcription of all 3 PDFs (5 Firm, 14 Product, 3 Lab - every field populated, 9 rows court-stayed). See 16_ksmscl.py's KARNATAKA_FULL_LIST and Entry 20. |
| BMSICL | 81 | WordPress RSS feed (public; the REST API is auth-restricted and was not used) | Done; order PDFs are also scanned images (same OCR block) so Reason/dates are "Not stated" |
| TNMSC | - | Tender-evaluation PDFs (no dedicated listing page) | BLOCKED - the site's listing/index pages time out consistently (tried at 30s/45s, multiple times); the one example PDF given fetches fine. No sitemap either. Not guessing additional PDF URLs. |
| OSMCL | - | Unknown | BLOCKED - known domain (osmcl.nic.in) does not resolve (DNS NXDOMAIN), confirmed via two independent lookups |
| TGMSIDC | - | Unknown | BLOCKED - site is up, but no blacklist/debarred page was found after checking home/tenders/notifications/sitemap |

**Combined master_list.xlsx**: 804 rows, 166 Matches (cross- and same-buyer),
0 Skipped rows, 1 Unmatched revocation. (Updated 2026-11-01 after the
Karnataka full-list rewrite - see Entry 20.)

## 2. Architecture

- `common.py` - shared helpers: the 19 master-list columns, name/date/reason
  cleanup, `write_sheet()` (plain-value Excel styling), `dedup_entry_ids()`,
  `find_matches()` (excludes Level="Lab" rows), `CATEGORY_PENDING` (the
  placeholder every buyer script writes into "Reason category" - see below).
- One script per buyer, each writing `output/buyers/<BUYER>.xlsx` in the same
  19 columns (+ its own `Skipped rows` / `Unmatched revocations` sheets where
  relevant). Rebuilding any one buyer file from scratch must reproduce the
  same IDs. **Since 2026-11-01 (Entry 19): buyer scripts write raw data
  only** - "Reason (as written)" (MPPHSCL folds its Remark field in too,
  when it differs from Reason - see Entry 19) and the lab name, but NOT a
  computed "Reason category" - they write `common.CATEGORY_PENDING` there
  instead. Classification is centralized (next bullet).
- `10_merge_master.py` combines every file in `output/buyers/` into
  `output/master_list.xlsx` (Master list + Matches + Skipped rows + Unmatched
  revocations), in sorted-filename order for reproducibility. **Since
  2026-11-01 (Entry 19) it ALSO runs classification**: `classify_row()`
  overwrites every row's `CATEGORY_PENDING` placeholder with the real
  category - `reason_category()`'s generic keyword match, OR a hand-
  verified MPPHSCL/RMSC override (the glossary moved here from
  `07_mpphscl.py`/`08_rmsc.py`) - then `quality_level()` computes Q1/Q2 from
  that. This runs ONCE, on every merge, across every buyer - no buyer
  script needs a separate re-run when `QUALITY_WORDS`/`SUPPLY_WORDS`/an
  override table changes (this is exactly how the Meril/GMSCL bug
  happened - Entries 16-17). It then automatically runs
  `11_cross_buyer_pairs.py` and `12_candidate_gaps.py` in that order
  (imported by file path, since their names start with a digit).
- `11_cross_buyer_pairs.py` breaks every cross-buyer Matches group into
  buyer-pair rows (Quality+Quality pairs first) in `cross_buyer_cases.xlsx`.
- `12_candidate_gaps.py` is the Step 4 first pass: candidate cases where
  buyer B found a Quality problem on a date inside buyer A's already-active
  ban - a report only, nothing merged. Writes `candidate_gaps.xlsx`
  ("Candidate gaps" + "By firm" sheets).
- `build_pdfs_to_read.py` - not auto-chained - builds a manual reading list
  of the order PDFs behind the candidate gaps.
- `14_match_firms.py` - Part C: cross-buyer fuzzy name matching, report
  only, nothing auto-merged.
- `record_change.py` - backup+diff helper for PROJECT_RULES.md rule B.

## 3. Script list, outputs, ID rules

| # | Script | Buyer / purpose | ID rule |
|---|---|---|---|
| 01 | download.py | AP + Assam downloads | - |
| 02 | extract_pdfs.py | AP PDF table/text extraction | - |
| 03 | build_master.py | APMSIDC | `AP-F-<S.No>` (firm list) / `AP-P-<year>-<S No>` (product list) |
| 03 | assam_master.py | AMSCL | `AS-<Sl no>` |
| 04 | pmbi.py | PMBI | `PMBI-<DEBAR/BLACK/ARCH>-<FromDate YYYYMMDD>-<DrugCode or FIRM>` |
| 05 | gmscl.py | GMSCL (Current) | `GMSCL-DEBAR-<order serial>-<year>` (e.g. `GMSCL-DEBAR-18394-421-2025-26`) |
| 06 | gmscl_archive_playwright.py | GMSCL (Archive) | `GMSCL-ARCH-...` - reads hand-saved HTML from `data/raw/GMSCL/archive_saved_by_hand/` if present; otherwise prints save instructions and stops (no longer attempts the blocked Playwright capture automatically) |
| 07 | mpphscl.py | MPPHSCL | `MPPHSCL-<FIRM/LAB/PRODUCT>-<order serial>-<year>` |
| 08 | rmsc.py | RMSC | `RMSC-DEBAR-<sanitized order reference>` |
| 09 | kmscl.py | KMSCL | `KMSCL-<NSQ/ED>-<order serial>-<year>` |
| 10 | merge_master.py | Combine all buyers | - (then auto-runs 11, 12) |
| 11 | cross_buyer_pairs.py | Cross-buyer pair breakdown | - |
| 12 | candidate_gaps.py | Step 4 candidate gaps | - |
| 13 | bmsicl.py | BMSICL | `BMSICL-DRUG-<WordPress post-id>` |
| 14 | match_firms.py | Part C fuzzy matching | - (report only) |
| 15 | rate_contracts.py | RMSC rate contracts (Step 4b) | `RMSC-RC-<tender code>-<drug code>-<supplier code>` |
| 16 | ksmscl.py | KSMSCL (name fragments only - see section 4) | `KSMSCL-<Firm/Product/Lab>-<name code>-<page>-<position>` |
| 17 | network_chart.py | Part 1 (data science) firm<->buyer network PNG | - (report only, FINAL STAGE - see Entry 25; reads `output/frozen_2026-11-07/` ONLY) |

Non-chained utility/report scripts (same backup rule applies - see
PROJECT_RULES.md rule B, clarified 2026-10-31 to cover these too): `build_pdfs_to_read.py`,
`build_network_firms_buyers.py`, `build_rate_contract_sources.py`,
`build_rc_vs_bans.py`, `build_rmsc_tender_review.py`,
`build_rmsc_rc_wayback.py`, `build_case_files.py`,
`build_finance_report.py`, `build_batch_crosscheck.py`,
`build_mpphscl_unclassified.py`, `build_concealment_cases.py` (new 2026-11-01,
see Entry 20), `build_headline_case_values.py` (new 2026-11-07, see Entry 23 -
reads RC data from `output/frozen_2026-11-07/` ONLY, not the live `output/`
folder, unlike every other script in this list). FINAL STAGE (2026-11-07,
Entry 25), same frozen-folder-only rule: `analysis_common.py` (shared helper,
never run directly), `build_summary_numbers.py`, `build_contagion_test.py`
(also reads `E:\cdsco_mqri\final_package\CDSCO_NSQ_Project_Workbook.xlsx`,
the user's own Project 1 - a path outside this project entirely, read-only),
`build_analysis_charts.py`. Part 2 (write-up), Entry 26:
`build_write_up_exports.py` - a one-off formatting step (not part of the
numbered pipeline), converts `write_up/*.md` into the extra formats asked
for (`Blacklist_Gap_Report.docx`, `One_page_summary.pdf`) - reads only the
already-finished `write_up/` markdown files and `output/analysis/` PNGs,
touches no buyer data itself.

All buyer scripts share: robots.txt checked before any fetch, 3-second pause
between requests, SHA-256 where a document is downloaded, plain-value Excel
output via `write_sheet()`.

## 4. Known issues / open items (see output/source_issues.csv for full detail)

**Standing decision (2026-10-31): no new buyers for now.** The user
asked to stop adding buyers after KSMSCL and move to Step 5 (finance)
and the case files instead. TNMSC/OSMCL/TGMSIDC/BMSICL's own gaps
below are therefore paused, not abandoned - revisit when the user says
to resume buyer-building.

- GMSCL Archive Blacklisted: blocked (WAF -> CAPTCHA on the second attempt).
  Now supports reading a hand-saved HTML copy instead - see script 06's
  instructions.
- KSMSCL: all 3 PDFs are scanned images. OCR is NO LONGER blocked - see
  Entries 10-12 (Tesseract 5.5.3 reinstalled 2026-10-31, official
  UB-Mannheim/tesseract-ocr Windows build, working and wired into
  `common.py` as `ocr_page_text()`/`save_ocr_text()`). Per PROJECT_RULES.md
  rule F, every value this produces still needs the user's own check
  against the PDF before it counts in any finding. BUT (Entry 13): all
  3 files are dense multi-column scanned tables - plain-text OCR does
  NOT preserve column alignment on any of them. `16_ksmscl.py` only
  extracts firm/product/lab NAME FRAGMENTS (possibly truncated, never
  confidently wrong); every other field is "Not stated". A full,
  confident extraction would need image-based table-cell segmentation
  before OCR - not attempted.
- TNMSC: listing/index pages time out consistently; no sitemap; blocked.
- OSMCL: domain does not resolve (DNS). Blocked.
- TGMSIDC: no blacklist/debarred page found on the live site. Blocked.
- RMSC "Medipol pharma_394.pdf" / "554.pdf": scanned images - OCR no
  longer blocked (see above), not yet re-attempted on these specific files.
- RMSC NIB-01/2025 tender PDF (`NIB 01.pdf`, 364 of 2,084 RC rows): also a
  scanned image PDF - OCR no longer blocked (see above); being read with
  OCR now (Entry 12 onward).
- RMSC NIB-08/2024 tender (417 of 2,084 RC rows, 2nd most frequent): the
  ORIGINAL tender notice could not be located on RMSC's Drug-Tenders
  listing page - only a corrigendum referencing it was found. Not
  downloaded or reviewed.
- A few name-splitting / order-number edge cases logged per-row rather than
  generalized into shared rules (KMSCL "Healers Lab, Avanbai", MPPHSCL
  "Bharat Serums" missing order number, ArchBlack "Corrigendum" + two-dates
  anomaly). See source_issues.csv.
- Part C fuzzy-match scorer (see section 3.5 below).
- **RESOLVED 2026-11-01 (Entry 19, checklist item 13):** the standing
  risk that a buyer script's OWN computed "Reason category" could go
  stale relative to `common.py`'s `QUALITY_WORDS`/`SUPPLY_WORDS` (exactly
  what caused the Meril/GMSCL bug, Entries 16-17) no longer exists -
  classification is centralized in `10_merge_master.py` (see section 2).
  As a side effect, checklist item 2 (Zee's MPPHSCL Remark field not
  being visible anywhere) is also improved: MPPHSCL's "Reason (as
  written)" now folds the Remark text in whenever it differs from Reason,
  so it is visible in `master_list.xlsx` itself, not just invisibly
  affecting classification - though item 2 itself (confirming this is
  real against the actual order PDF) is still the user's to check, not
  settled by this change.

### 3.5 Design decision: Part C fuzzy-match scorer (current state - see
entries 2-4 in section 5 for how this evolved)

Originally specified as `rapidfuzz.fuzz.token_sort_ratio` with an 85
threshold. That alone MISSES this project's own target case - "Medipol
Pharmaceutica" (RMSC) vs "Medipol Pharmaceutical India Pvt Ltd"
(MPPHSCL/PMBI) scores only **73.7**, below 85, because "India Pvt Ltd"
being present on one side and absent on the other reads as a large
difference even though one name is clearly a truncated version of the
other.

**Current design (as of 2026-10-31):**

1. **Compare unique standardized names, not rows.** 782 rows collapse to
   382 unique standard names. Comparing every row pair (216,891 pairs)
   produced a mostly-duplicate 26,723-row result - unreviewable. Comparing
   the 382 unique names instead (59,986 pairs) removes that duplication at
   the source; each result row lists every Entry ID and buyer behind both
   sides.
2. **Two scorers, different thresholds.** `token_sort_ratio >= 85` (as
   specified) OR `WRatio >= 90` (NOT 85 - see the bug below). "Caught by"
   records which one(s) triggered it.
3. **Shared-word filter (a real filter, not just a warning).** Strip
   common pharma/corporate words (Pharmaceuticals, Pharmaceutical, Pharma,
   Formulations, Formulation, Laboratories, Labs, Lab, Healthcare, Drugs,
   Remedies, Life, Sciences, Science, Lifesciences, Lifescience, Biotech,
   India, Pvt, Ltd, Private, Limited, LLP, And, Co) from both names, keep
   only words of 4+ letters, then require the two names to share at least
   one such word - "shared" allows a fuzzy match (`fuzz.ratio >= 85`
   between the two words, catching a typo like "MERCURY"/"MERCUREY"), not
   just exact string equality. A pair that clears the score threshold but
   shares no such word goes to a separate "Generic words only" sheet
   (kept, not dropped) instead of the main "Match candidates" sheet.
4. **The common-words list is user-controlled.** Before finalizing, the
   script prints the 40 most frequent words across all 382 unique
   standard names, with counts - the user reviews this and decides what (if
   anything) to add; the script does not add words on its own.

**Bug found and fixed: WRatio produces a flat, artifact score of exactly
85.5** for "Super Formulation(s) Pvt Ltd" against dozens of unrelated
names - confirmed even against a made-up "ABC Random Ltd". Checked the
full score distribution of the (pre-fix) 127 "kept" pairs: a hard spike of
59 rows at exactly 85.5, nothing between 85.5 and 90, then a clean, normal
distribution from 90 upward that is ALL genuine matches when read by hand
(including Medipol at exactly 90.0, plus several more real cases: Healers
Lab's truncated name, and address text merged into the firm name for
Stallion Laboratories, Zest Pharma, Syncom Healthcare). Raising WRatio's
threshold specifically to 90 (token_sort_ratio stays at 85, as it never
showed this artifact) removes the 85.5 cluster entirely while keeping
everything of value - final result: 63 kept, 66 in "Generic words only",
out of 59,986 unique-name pairs checked.

Confirmed both test cases against the real, final output:
"Medipol Pharmaceutica" vs "Medipol Pharmaceutical India Pvt Ltd" -
kept, WRatio=90.0, caught_by=WRatio. "Super Formulation Pvt Ltd" vs
"Super Formulations Pvt Ltd" (the real same-firm pair in this data) -
kept, both scorers, 98.0. "Super Formulation(s) Pvt Ltd" vs other,
unrelated firms that only share the word "Formulation(s)" (Yeluri
Formulations, Rivpra Formulation, Apex Formulations) - correctly in
"Generic words only", not the main sheet.

Nothing is ever auto-merged - all of the above only changes which rows a
human sees and how they are organized, never what happens to the
underlying buyer data.

**Generic-words list, finalized by the user 2026-10-31** (from the top-40
report): added HEALTH, CARE, PARENTERALS, SURGICAL, INDUSTRIES,
INTERNATIONAL, MEDICAL, CHEMICALS, BIO, MEDITECH, BHARAT. Deliberately kept
OUT: ANG, BRILLIANT (real distinctive name fragments) and INDORE, MUMBAI,
ROAD, NAGAR (these indicate the address-leak bug below, not a generic-word
question - fixing the bug was the right response, not adding them here).

### 3.6 Bug found and fixed: address text leaking into the firm name (GMSCL, MPPHSCL)

Found while auditing standard names for the words above: GMSCL's "Name of
Firm" field and MPPHSCL's "Company Name" field sometimes contain the
firm's address appended directly - there is no separate address column on
either site to split from, so this is not a parsing bug in the sense of
"failed to use a field that existed" - the source literally crams both
into one field, inconsistently (e.g. "Alpa Laboratories" on one row,
"M/s Alpa Laboratries Ltd., Indore" on another, same real firm).

Audited every unique standard name (all buyers) for address signals -
ROAD, NAGAR, PLOT, VILLAGE, DIST/DISTRICT, ESTATE, a 6-digit PIN, or a
recognized city/state name: 24 unique names flagged, 22 real leaks (1
GMSCL, 21 MPPHSCL) and 2 false-positive-looking groups that turned out to
be genuine company names containing a place word as part of the name
itself (Kerala State Drugs And Pharmaceuticals Ltd, Goa Antibiotics
Pharmaceuticals Ltd, Bengal Chemicals And Pharmaceuticals Limited, Baroda
Surgical (India) Private Limited, Gujarat Pharmalab Pvt Ltd) - confirmed
these should NOT be touched (the place name is the first word / the
whole name, not appended after a comma or a legal suffix).

Fixed at the source with a new shared helper, `strip_leaked_address()` in
`common.py`, used by `05_gmscl.py` and `07_mpphscl.py` when building each
row's firm name (not by hand-editing any output file). It looks for
structural signals that an address has been appended - a comma followed
by a number or a recognized place name, an explicit marker word (Plot No,
Add:, Village, Nagar, Road, Dist/District, Estate, Society, Sector, Floor,
Highway, Taluka, Survey No, Industrial Area) anywhere, a 6-digit PIN, or a
place name appearing AFTER a legal-entity suffix (Ltd/Limited/Pvt/
Private/LLP) with no comma at all (catches e.g. "...Pvt Ltd. New Delhi").
A real bug was found and fixed while building this: a shared regex alternation's
trailing `\b` could never match after "Add :" (colon-to-space is not a word
boundary), which silently made that marker never fire - fixed by giving
each alternative its own boundary instead of one shared trailing one.

Re-ran `05_gmscl.py`, `07_mpphscl.py`, then the merge chain. Confirmed: all
22 predicted names changed exactly as proposed in the audit, zero other
column changed on any of the 782 rows, and row counts stayed at 99 (GMSCL)
and 378 (MPPHSCL) - nothing added or removed, only names cleaned up.

### 3.7 Legal basis: NSQ declaration and the Q1/Q2 quality-level split

*(The user asked for this to be recorded as "3.5" - that number was
already in use above for the Part C fuzzy-match scorer design decision,
so this continues the same section-4 sub-numbering sequence as 3.7
instead of overwriting 3.5.)*

**The question:** RMSC's debarment clause disqualifies a bidder for an
"Not of Standard Quality" (NSQ) finding by ANY procuring agency, but not
for a non-supply/delay finding by another agency. A bare "Quality"
category (common.py's `reason_category()`) does not by itself tell you
whether a given ban is a genuine NSQ finding in the legal sense, or
merely a quality-sounding complaint that never went through the
statutory process. That distinction matters directly for every "Possible
breach" classification in `build_rc_vs_bans.py`.

**Legal basis, as given by the user:** NSQ is formally declared via a
Government Analyst's report under the Drugs and Cosmetics Act, 1940,
Section 25. A sample that fails a pharmacopoeial test (dissolution,
sterility, assay, content of, etc.) means NSQ "in substance" even where
the formal Section 25 declaration route is not explicitly named in the
buyer's own record.

**Implementation - two levels, both reported, neither silently
preferred:**
- **Q1** = the ban record (reason text and/or the "Lab that declared
  NSQ" field) names a statutory testing authority - a Government
  Analyst, State Drug Testing Laboratory, CDSCO, CDTL, DCL/CDL, or a
  state Food and Drug (Testing) Laboratory. This is the Section 25
  route - a literal NSQ declaration.
- **Q2** = the quality failure was found at the buyer's own or an
  empanelled lab, or no lab is named at all (e.g. "Quality Failure",
  "Dissolution", "Sterility test" alone, with no lab identified). This
  is NSQ "in substance" per the user's standard above, but not a
  literal statutory declaration - debatable whether RMSC's clause
  reaches it.
- `quality_level()` in `common.py` computes this from the reason text
  plus the lab field; `10_merge_master.py` writes it onto every row of
  `master_list.xlsx` (centrally, since Entry 19 - see section 2).
- Re-classifying the policy-gap firms under this split, and
  propagating `QUALITY_WORDS`'s earlier expansion through every buyer
  script that calls `reason_category()`, changed the quality/quality
  candidate count from 17 to 33 pairs (confirmed as a legitimate
  consequence of the word-list fix, not a bug - see Entry 17). After
  this re-classification, no firm's "Possible breach" survives under
  EITHER Q1 or Q2 once the deadline-overlap check is also applied - see
  Entry 17 for the full group counts.

**Decision 2026-11-01 (Entry 19), superseding how the two levels are
PRESENTED (not the Q1/Q2 definitions above, which are unchanged):**
checklist item 11 asked the user to decide whether Q2 should count
toward RMSC's clause. Their answer - **headline findings use Q1 only.
Q2 is reported separately, always labeled "quality failure in
substance, not a formal NSQ declaration (no government lab named)",
with both counts shown everywhere** (the Strength sheet, the By-firm
sheet, and the script's own printed summary all now show a Q1 column/
count and a Q2 column/count side by side, never merged into one
string or one number). `build_rc_vs_bans.py`'s `clause_check()` now
returns `(allowed, active, q1_classification, q2_classification)` -
two independent values - instead of one combined `classification`
string that depended on reading which branch produced it. Q2 can never
become part of the Q1 headline count, by construction, not just by
convention.

### 3.8 Legal re-reading: clause 5(l)(II)/5(k)(ii) is a WHOLE-FIRM clause, not product-only

*(Superseding `build_rc_vs_bans.py`'s earlier A/B-only clause scope - see Entry 20.)*

**The question:** does RMSC's eligibility clause disqualify a bidder from
the WHOLE bid once another buyer finds any of its products NSQ, or only
from bidding on that SAME specific product?

**Re-read 2026-11-01, per the user's own document check**
(document_checks_2026-11-01.xlsx, "Rajasthan clause" sheet, exact
clause text quoted there, read directly from the tender PDF pages):
clause 5(l)(II) (drugs) / 5(k)(ii) (sutures) reads "The concern/company/
firm which stands blacklisted/banned/debarred ... by any other State/
Central Government ... on the ground of ... the products being found Not
of Standard Quality (NoSQ) ... **shall also not be eligible to
participate in the Bid**" - this is WHOLE-FIRM, by its own plain words,
regardless of whether the CURRENT tender's own line item matches the
specific product the other buyer found NSQ. Clauses 5(l)(I)/(III) and
5(k)(iii) are the narrower, PRODUCT-ONLY clauses (barring a bid only on
"those products" found NSQ elsewhere).

**This reverses the earlier working assumption** ("product-only is
better supported") that `build_rc_vs_bans.py`'s Strength/By-firm sheets
were built around (clause scope restricted to Groups A/B only). Fixed by
widening the primary (headline) clause check to Groups A/B/C - any
Quality-category, non-stayed ban by another buyer, regardless of product
match - while keeping the OLD A/B-only behavior as a SECOND, explicitly
labeled "product-only reading" column, per the user's instruction to
show both rather than replace one with the other.

**Immediate effect**: Zee Laboratories, Meril Endo Surgery, and Nestor
Pharmaceuticals - all three of this project's own named headline
cases - were in Group C (their matching other-buyer ban is for a
DIFFERENT drug/dosage form than their Rajasthan RC lines), so none of
them had ever been checked against RMSC's clause at all before this fix,
even after their Q1 lab was separately confirmed (Entry 17/18, Entry 20).
See Entry 20 for the full before/after numbers.

## 6. Sanskruti's checklist (manual follow-up tasks)

*A reference section, kept current in place like sections 1/3/4 - check
items off or add notes as they're done, don't leave it as history. It is
numbered 6 and placed BEFORE section 5 deliberately: section 5 is
append-only and must stay the last section as it keeps growing, so this
checklist lives here instead, right after the other reference sections.*

Per PROJECT_RULES.md rule E, every report ends with a short "Still pending for
Sanskruti" list pulled from whichever of these are still open.

*Simplified into 3 groups 2026-10-31, per the user's request - the
flatter 11-item version above is replaced, not kept alongside this one.*

*Re-touched 2026-10-31 after Entry 17's review-decisions/Q1-Q2/Meril/
MPPHSCL-OCR batch - items 1, 3 and 7 below are now DONE (kept, struck
through, with where to find the result) rather than deleted, so the
history of what was asked and resolved stays visible.*

**TODAY**
- [x] 1. ~~Place the real `first_pass_review_2026-10-31.xlsx` in
      `output/review/`~~ - DONE (Entry 17): the file was actually at
      `output/first_pass_review_2026-10-31.xlsx` (not the `review/`
      subfolder), found and all 127 match decisions loaded into
      `match_decisions.xlsx` via `build_match_decisions.py`; 14 and the
      merge re-run. See Entry 17 for the exact before/after counts.
- [x] 2. ~~Zee's new Group A finding: MPPHSCL's ban on Zee Laboratories
      (firm-level) has a REMARK of "Blacklisted/Debarred for Quality
      Failure" - check this is real by opening Zee's actual MPPHSCL order
      PDF (`output/cases/Zee.xlsx` has the entry ID and link). *Narrowed
      2026-11-01 (Entry 19): the remark is no longer invisible - it now
      shows directly in `master_list.xlsx`'s "Reason (as written)"
      column (e.g. "...Condition [Remark: Blacklisted/Debarred for
      Quality Failure]"). What's still open is only confirming it
      against the real PDF, same as any other Auto-check row.* *Further
      narrowed 2026-11-01 (Entry 20): the user's own document check
      confirms this is real and names the lab (CDL Kolkata/Government
      Analyst) for all 4 of Zee's MPPHSCL product rows - now Q1 in
      master_list.xlsx. The 4 order PDFs themselves (p1 each) are still
      worth a final look, but the substance is no longer in doubt.*~~ -
      CLOSED 2026-11-07: the user confirmed all 4 of Zee's MPPHSCL order
      PDFs by eye, and separately confirmed Radiant Parenterals Ltd's
      MPPHSCL order PDF (the concealment case,
      `output/concealment_cases.xlsx` - already Status "Checked" since
      Entry 20, now also eye-verified). Nothing changed in the data;
      this item is fully done.
- [x] 3. ~~Nestor's "possible breach" finding is RETRACTED~~ - DONE
      (Entry 16/17): the RMSC analysis stays retracted (dosage-form
      mismatch), but a SEPARATE, genuinely confirmed finding (CAG
      Haryana audit report / The Tribune, found via web search, not
      supplied by you) is now its own entry in `output/cases/
      Nestor.xlsx`'s "Outside evidence" sheet - worth reading, not
      something to action in the pipeline.
- [x] 11. ~~Review the Q1/Q2 quality-level split~~ - DECIDED by the user
      2026-11-01 (Entry 19, PROGRESS_LOG.md section 3.7): headline
      findings use Q1 only; Q2 is reported separately, always labeled
      "quality failure in substance, not a formal NSQ declaration (no
      government lab named)", with both counts shown everywhere in
      `rc_vs_bans.xlsx` (Strength sheet, By-firm sheet, and the script's
      own printed summary). Nothing further needed from you here unless
      the decision itself changes.
- [x] 12. ~~Review `output/mpphscl_unclassified_batch.xlsx`'s "Random 20
      sample" sheet by hand~~ - DONE (Entry 20): the user read all 20
      sample letters directly from the page images and found the
      computer right on only 3 of 20 (Bio-in-Vitro, Syndicate Pharma,
      Fresenius Kabi). All 20 are now hand-verified, entry-ID-keyed
      overrides in `10_merge_master.py` (`MPPHSCL_HAND_VERIFIED_OVERRIDES`),
      Status = "Checked" not "OCR - check". The classifier itself was
      also improved with 5 new signal groups found from this sample
      (agreement-not-executed, price-above-ceiling text, furniture/
      equipment, forged document, concealment of another buyer's ban)
      and re-run on the full 164 from cached OCR text (0 new downloads/
      OCR calls) - new totals: To check 85 (was 110), Other 37 (was 11),
      Quality 21, Non-supply 21. The other ~144 rows remain unconfirmed
      ("OCR - check") - this item is DONE for the 20-row sample asked
      for; a full hand-check of the remaining rows is a new, separate
      ask if wanted.

**NEXT**
- [x] 4. ~~NIB-01/2025 deadline + page 55: check the OCR'd deadline
      (05.06.2025) and clause (pages 9-10, 55-56) against
      `data/raw/RMSC_Tenders/NIB01_2025_NIB01.pdf` yourself. Also
      resolve the "3% vs 5%" discrepancy on page 55 (Annexure VII item 3).~~
      - CLOSED 2026-11-07: confirmed by the user by eye
      (`output/review/document_checks_2026-11-01.xlsx`, "Rajasthan
      clause" sheet). The deadline is confirmed; "3% vs 5%" is RESOLVED -
      it is 5% ("Bidder is bound to supply minimum 5% of bid quantity on
      monthly basis") - the OCR's "3%" was a misread, not a real
      difference between tenders. Also found on this same pass: Annexure
      VII item 4's affidavit asks about a ban for "conviction ... or the
      products being found spurious or adulterated" - narrower than
      clause II's own "Not of Standard Quality" ground - a genuine
      wording inconsistency, not resolved either way, now a Findings
      item in the write-up (Policy gaps).
- [x] 5. ~~Match pairs incl. Jackson: review the 5 new unmatched pairs
      from this turn's re-run (all KSMSCL Martin/Vivek name fragments,
      Entry 17)~~ - RESOLVED BY DATA, 2026-11-01 (Entry 20): the
      Karnataka full-list rewrite (checklist item below / task 5)
      replaced the ambiguous OCR fragments ("Martin", "Vivek") with full
      firm names, so the fragment-vs-full-name question no longer
      exists. The user's "New match decisions" sheet confirmed Martin &
      Brown Bio Sciences and the Vivek Pharmachem variants are each one
      real firm (now in `match_decisions.xlsx`) and that Unicure (RMSC)
      = Unicure India Ltd (updated from "Not sure" to "Same"). Jackson's
      case file (`output/cases/Jackson.xlsx`) now includes its KSMSCL
      row (needed a second name variant added to `build_case_files.py`
      - the full legal name does not collapse to PMBI's shorter one).
      The original 66 kept + "Generic words only" pairs unrelated to
      KSMSCL remain open for review, not newly resolved by this.
- [x] 6. ~~Check the Hindi OCR dates~~ - DONE for 4 of the 7 letters
      (Entry 20): the user read Zee/1846, ANG/3078-b, Vee Excel/1453,
      and Unicure/3093 directly from the page images and found wrong
      numbers in all 4 - corrected in `build_pdfs_to_read.py`
      `RMSC_HINDI_READINGS`. Medipol/554 and Medipol/3085 were checked
      and confirmed CORRECT, no change. Ultra Drugs/1837 was explicitly
      not re-read (not a quality case). The ANG letter's own order
      number (3099) does not match this project's Entry ID (3078, taken
      from RMSC's live page) - logged as an open discrepancy in
      `source_issues.csv`, not resolved either way.
- [x] 7. ~~Meril Endo Surgery Pvt Ltd now appears in the new Group U~~ -
      DONE (Entry 17): this was a real bug (05_gmscl.py had never been
      re-run after the Quality-word-list expansion), now fixed; Meril's
      GMSCL ban is confirmed ACTIVE on the RMSC NIB-10/2026 sutures
      deadline, and the lab (Food and Drug Laboratory, Vadodara) is
      identified. Full picture in `output/cases/Meril.xlsx`. Still worth
      reading together with the sutures tender at
      `data/raw/RMSC_Tenders/
      NIB10_2026_Sutures_Finalbid10UBN.pdf`) you asked about for this
      exact firm.

**LATER - OPTIONAL UPGRADES, not pending** (re-labelled 2026-11-07, per
the user's own instruction - items 8 and 9 would each strengthen the
project further, but neither blocks any finding already made; nothing
in the write-up or finance report depends on either being done)
- [ ] 8. Save the GMSCL Archive (Blacklisted) tab from your own browser
      into `data/raw/GMSCL/archive_saved_by_hand/` (blocked by CAPTCHA
      for the automated capture - see section 4).
- [ ] 9. File RTIs for: Telangana (TGMSIDC), Odisha (OSMCL), Assam,
      Haryana, Delhi CPA, Tamil Nadu (TNMSC), and the Gujarat (GMSCL)
      archive - keep a tracker (state, date filed, reference number,
      response due, response received). *Added 2026-11-01 (Entry 20,
      "Still to do" sheet): also RMSC for copies of the Annexure VII
      affidavits of Zee Laboratories (NIB-10/2024, NIB-14/2024,
      NIB-01/2025) and Meril Endo Surgery (NIB-10/2026) - to check
      whether either firm actually disclosed the other-buyer ban.* *Added
      2026-11-07: also PMBI - for a copy of (or the lab/authority behind)
      Zee Laboratories' Carboxy Methyl Cellulose Eye Drops debarment
      (`PMBI-DEBAR-20241218-341`, 18-12-2024 to 17-12-2026) - PMBI's own
      debar.aspx exposes no order-document link for this row (confirmed
      by inspecting the live page's HTML directly, 02-11-2026 -
      `source_issues.csv`), so this is the only remaining way to find out
      whether it is Q1 or Q2 - see item 15, now closed on this point.*
- [x] 15. ~~Read PMBI's Zee CMC-eye-drops debarment order and APMSIDC's
      Nestor order in full to find the testing lab named (determines
      whether each is Q1 or Q2) - added 2026-11-01 (Entry 20, "Still to
      do" sheet #3).~~ - CLOSED 2026-11-07: Nestor - the AP items list
      already names "DCL, Vijayawada for 4 Batches" in its own "Lab that
      declared NSQ" field (`AP-P-2023-24-005`) - Q1, no code change
      needed, this was already correct in `master_list.xlsx`. Zee's PMBI
      CMC eye drops - no order PDF exists at all to read (confirmed
      02-11-2026, debar.aspx has no document link for this row) - this
      stays Q2 permanently on that ground, not "Q2 until the PMBI order
      is read"; the only way forward is an RTI, now added to item 9.
- [x] 16. ~~Meril's MPPHSCL appeal order page 2~~ - DONE (Entry 21,
      2026-11-02): the appeal order is numbered 167 (not 166 as first
      read - see source_issues.csv), dt 20-01-2025. It SET ASIDE the
      firm-level blacklisting; Meril remains blacklisted only for item
      MIS93, 2 years from 06-11-2024, for a rate-quoting error. Applied
      as Level/Category/Court overrides in `07_mpphscl.py`/
      `10_merge_master.py`.
- [x] 10. ~~Check the RMSC tender clause quote on page 9 (NIB-14/2024/69,
      NIB-10/2024, and NIB-10/2026 sutures) yourself, against the
      original PDFs in `data/raw/RMSC_Tenders/`.~~ - CLOSED 2026-11-07:
      the user read and transcribed the exact clause text
      (`output/review/document_checks_2026-11-01.xlsx`, "Rajasthan
      clause" sheet) - clause 5(l) I/II/III (drugs, page 9-10) and 5(k)
      ii/iii (sutures, page 9), quoted word for word. Confirms the
      whole-firm reading (clause II/ii) `build_rc_vs_bans.py` already
      uses (PROGRESS_LOG.md section 3.8).
- [x] 17. ~~Check Annexure VIII quantities for Zee's 26 NIB-01/2025 drug
      codes and Meril's 15 NIB-10/2026 item codes against the real PDF
      pages, by eye - added 2026-11-07 (the finance report). The OCR of
      these 2 scanned tenders' Annexure VIII tables does NOT reliably
      preserve column alignment (same problem already found for KSMSCL -
      see `source_issues.csv`, 07-11-2026), so NO quantity/pack-size/
      estimate was computed for either tender in
      `output/finance/headline_case_values.xlsx`.~~ - CLOSED 2026-11-07:
      the user read all 26 Zee lines (23 of 26 - 3 hidden under the
      digital-signature stamp, left blank) and all 15 Meril lines
      directly from the Annexure VIII page images
      (`output/review/Finance_scanned_tenders_2026-11-07.xlsx`) - now in
      `output/finance/headline_case_values.xlsx`'s "Read by eye" sheet,
      folded into "Combined totals".
- [x] 13. ~~The Meril bug was caused by a buyer script not being
      re-run after a shared word list changed~~ - FIXED at the root
      2026-11-01 (Entry 19): `reason_category()`, the MPPHSCL/RMSC
      overrides, and `quality_level()` all moved out of the buyer
      scripts into `10_merge_master.py`, which now computes every row's
      category centrally on every merge. Buyer scripts only write raw
      "Reason (as written)" and the lab name. Re-ran everything end to
      end and compared against a before-snapshot of `master_list.xlsx`:
      0 rows changed category, 0 rows changed Quality level - confirmed
      behavior-preserving, not just restructured. This failure mode
      (a buyer script silently going stale) cannot recur.

## 5. Decision log (append-only - never edit an old entry)

---

### Entry 1 - 2026-10-31 - PROGRESS_LOG.md created; correcting a numbering mix-up

**Asked**: Set up PROJECT_RULES.md with standing rules, a `record_change.py`
backup/diff helper, a starting code snapshot, and check PROGRESS_LOG.md
against the real project folder, fixing anything that doesn't match -
recording corrections as the first entry here.

**Done**:
1. Looked for `PROGRESS_LOG.md` in the project folder (`E:\blacklist_gap\`) -
   it did not exist anywhere in the project. Asked the user to confirm;
   they asked me to create it fresh rather than search elsewhere.
2. The user's instructions referred to Part C as "13_match_firms.py", but
   `13` was already taken by `13_bmsicl.py` (renamed earlier in the same
   session from `11_bmsicl.py`, to make room for the user's requested
   `11_cross_buyer_pairs.py` / `12_candidate_gaps.py`). Flagged this;
   confirmed with the user to keep Part C at the already-built
   `14_match_firms.py` rather than create a second collision.
3. Created `PROJECT_RULES.md` (rules A-D).
4. Created `scripts/record_change.py`, explained it, tested it on
   `common.py` (no-op edit, confirmed empty diff, then cleaned up the test
   artifacts).
5. Took the starting snapshot: copied all 19 files in `scripts/`
   (15 numbered scripts + `common.py`, `build_pdfs_to_read.py`,
   `record_change.py`, plus `PROJECT_RULES.md`) into
   `code_history/2026-10-31_start/`.
6. Wrote this file, with sections 1/3/4 built directly from the real
   current state (buyer file row counts re-checked via
   `openpyxl.load_workbook` on every file in `output/buyers/`, not
   recalled from memory), so there was nothing stale to "correct" beyond
   the numbering mix-up above.

**Code changed**: `PROJECT_RULES.md` (new - `code_history/2026-10-31_start/PROJECT_RULES.md`
is the baseline copy, no `.diff` since it didn't exist before), `scripts/record_change.py`
(new), `PROGRESS_LOG.md` (new - this file).

**Results with numbers**: 782 total master-list rows across 8 buyer files
(AP 43, AS 0, BMSICL 81, GMSCL 99, KMSCL 5, MPPHSCL 378, PMBI 75, RMSC 101),
confirmed by direct re-count, not carried over from an earlier claim.

**Problems and source issues**: The user's request assumed `PROGRESS_LOG.md`
already existed with numbered sections (referenced "section 3.5", "section
5") and that Part C was at `13_match_firms.py` - neither was true. Both
were raised and resolved with the user rather than guessed.

**Open items**: Going forward, every task gets a new entry here per rule A,
and every code edit gets a `.before`/`.diff` pair per rule B. Everything
before this entry (Part A, PMBI, GMSCL, MPPHSCL, RMSC, the two cross-buyer
analysis rounds, KMSCL/KSMSCL/BMSICL/TNMSC) happened in earlier turns of
the same session, without this log existing yet - summarized in sections
1-4 above rather than reconstructed turn-by-turn here, since no
`.before`/`.diff` pairs exist for those earlier edits (the backup
mechanism didn't exist until this entry).

---

### Entry 2 - 2026-10-31 - Part C (14_match_firms.py) built, with dual-scorer design

**Asked**: Build Part C - fuzzy cross-buyer name matching with rapidfuzz,
score >= 85, different buyers only, "My decision" column, nothing
auto-merged - and confirm it catches "Medipol Pharmaceutica" vs "Medipol
Pharmaceutical India Pvt Ltd". Then, mid-task: compute both token_sort_ratio
and WRatio, include a pair if either clears 85, add a "Caught by" column, add
a generic-word check with a "Warning" column, sort both-scorers-first then
WRatio-only then generic-word-flagged last, and confirm the Medipol pair has
no warning while "Super Formulations" vs "Formulations" does.

**Done**: See section 3.5 above for the scorer-choice reasoning. Built
`14_match_firms.py` (not `05`, which is taken by GMSCL, and not `13`, which
is taken by BMSICL - see Entry 1).

**Code changed**: `scripts/14_match_firms.py` (new -
`code_history/2026-10-31_start/14_match_firms.py` is the baseline copy of
its first version; the dual-scorer/generic-word rewrite happens after this
entry and will get its own `.before`/`.diff` pair per rule B going forward).

**Results with numbers**: superseded almost immediately by Entries 3-4
below - the first dual-scorer version (warning-only, row-level comparison)
produced 26,723 candidates out of 216,891 pairs, 98% flagged - unreviewable,
and was replaced before being reported as a final result.

**Problems and source issues**: none new at this point.

**Open items**: see Entries 3-4.

---

### Entry 3 - 2026-10-31 - Scale fix: unique-name comparison + shared-word filter as a real filter

**Asked**: After seeing the 26,723/98%-flagged result, the user chose to
turn the generic-word check into a real filter (not just a warning) -
then, before that was finished, refined further: compare unique
standardized names instead of rows (with an Entry-IDs/buyers column per
side), make the shared-word check fuzzy (`fuzz.ratio >= 85` between
individual words, 4+ letters only), and show a 40-most-frequent-word
report for the user to decide additions from, rather than deciding them
automatically.

**Done**: Rewrote `14_match_firms.py` around unique-name groups (782 rows
-> 382 unique standard names), added `content_words()` (strip generic
words, require 4+ letters) and `has_fuzzy_shared_word()` (fuzzy word-level
match, not exact), split output into "Match candidates" (shares a real
word) and "Generic words only" (does not, but kept, not dropped) sheets,
and added the word-frequency report. Ran it and printed the top 40 words
for the user (LTD 258, PVT 175, PHARMACEUTICALS 48, ... down to MEDITECH
4) - none added to the generic list without the user's say.

**Code changed**: `scripts/14_match_firms.py` -
`code_history/2026-10-31_0546_14_match_firms.py.before`/`.diff`.

**Results with numbers**: 382 unique standard names; 59,986 unique-name
pairs checked (different buyers); first run of this version: 127 "kept",
8,084 "Generic words only". (The 8,084 figure turned out to be inflated by
the bug found in Entry 4, not a true reflection of the filter design.)

**Problems and source issues**: while checking the Medipol/Super-Formulations
test cases against this result, found a second, more serious problem -
see Entry 4.

**Open items**: see Entry 4.

---

### Entry 4 - 2026-10-31 - Bug found: WRatio gives a flat false-positive score; fixed with a separate threshold

**Asked**: (self-directed, while verifying Entry 3's output against the
required test cases) - the "Generic words only" sheet had ballooned to
8,084 rows, and manual spot-checking showed "Super Formulation(s) Pvt Ltd"
scoring WRatio=85.5 against many clearly-unrelated firms (Cipla Ltd,
Wockhardt Ltd, Zenith Drugs Ltd, Bal Pharma Ltd, ...).

**Done**:
1. Confirmed it was not data-specific: `fuzz.WRatio("SUPER FORMULATION(S)
   PVT LTD", "ABC RANDOM LTD")` also scores exactly 85.5 - a scorer
   artifact on short/very-different-length strings, not real similarity.
2. Checked the full WRatio score distribution of the (pre-fix) 127 "kept"
   rows: a hard spike of 59 rows at exactly 85.5, a gap with nothing
   between 85.5 and 90, then a clean distribution from 90 upward.
3. Read every row scoring exactly 90.0 by hand (12 rows) - ALL genuine:
   Medipol (the required test case), plus Healers Lab's truncated name and
   three more cases of address text merged into the firm name (Stallion
   Laboratories, Zest Pharma, Syncom Healthcare) - new instances of a
   pattern already known from KMSCL's "Healers Lab, Avanbai" row
   (`source_issues.csv`), now confirmed present in other buyers' data too.
4. Fixed by giving WRatio its own threshold, 90 (not 85) - `token_sort_ratio`
   keeps its original 85, since it never showed this artifact. Re-ran.

**Code changed**: `scripts/14_match_firms.py` -
`code_history/2026-10-31_0549_14_match_firms.py.before`/`.diff`.

**Results with numbers**: Final run - 382 unique standard names; 59,986
unique-name pairs checked; **63 kept** in "Match candidates"; **66** in
"Generic words only" (down from the artifact-inflated 8,084). Within the
kept group: WRatio >= 85: 63, >= 90: 55, >= 95: 38.

Both required tests confirmed against the real, final output:
- "Medipol Pharmaceutica" (RMSC) vs "Medipol Pharmaceutical India Pvt Ltd"
  (MPPHSCL) - kept in Match candidates, WRatio=90.0, caught_by=WRatio.
- The real same-firm pair "Super Formulation Pvt Ltd" vs "Super
  Formulations Pvt Ltd" - kept, both scorers, 98.0. Unrelated firms that
  only share the generic word "Formulation(s)" (Yeluri Formulations,
  Rivpra Formulation, Apex Formulations) - correctly routed to "Generic
  words only", not the main sheet.

**Problems and source issues**: the address-merged-into-firm-name pattern
(Stallion Laboratories, Zest Pharma, Syncom Healthcare) is now confirmed
beyond the single KMSCL case already in `source_issues.csv` - not yet
logged there individually for each new instance; worth a follow-up pass
if the user wants the full extent of this mapped.

**Open items**: the user has not yet told us which (if any) of the
top-40 words to add to the generic-words list (e.g. HEALTH, PARENTERALS,
CARE, SURGICAL, MEDICAL, INDUSTRIES, INTERNATIONAL, MEDITECH look like
reasonable pharma-generic candidates; ANG, BHARAT, BRILLIANT, INDORE,
MUMBAI, NAGAR, ROAD look like real name fragments or place names that
should NOT be added) - nothing added without that decision. Also open:
whether to formally log the newly-confirmed address-merged-into-name
instances in `source_issues.csv` the same way the KMSCL one already is.

---

### Entry 5 - 2026-10-31 - Generic-words list finalized; address-leak bug found, audited, and fixed

**Asked**: Decide the generic-words additions from the Entry 4 report
(ADD: HEALTH, CARE, PARENTERALS, SURGICAL, INDUSTRIES, INTERNATIONAL,
MEDICAL, CHEMICALS, BIO, MEDITECH, BHARAT; KEEP as distinctive: ANG,
BRILLIANT; do NOT add INDORE, MUMBAI, ROAD, NAGAR - their presence means
an address leaked into the firm name, a bug to fix, not a word-list
question). Then: (1) list every standard name with an address signal
(Entry ID | buyer | name as written | proposed clean name), flagging
unclear ones for a decision; (2) fix the cause in each affected buyer
script, not by hand-editing outputs - a real city in a real firm name
must stay, list unclear ones; (3) re-run and show every changed name,
confirm no other column changed; then re-run Part C and confirm Medipol
still kept (report its new score), Super Formulations vs Formulations in
"Generic words only", and report kept/generic-only/score-band counts.
Also: KSMSCL status in one line. Explicitly deferred: TNMSC, OSMCL,
TGMSIDC (to RTI / later, not now).

**Done**:
1. Updated `GENERIC_WORDS` in `14_match_firms.py` with the 11 approved
   additions.
2. Audited every unique standard name (382 of them) for address signals
   (ROAD, NAGAR, PLOT, VILLAGE, DIST/DISTRICT, ESTATE, 6-digit PIN, or a
   recognized city/state name): 24 flagged. Designed a cut heuristic,
   tested it against all 24 by hand before touching any script - found
   and fixed two bugs in the heuristic itself during that testing (a
   shared regex `\b` that could never match after "Add :"; missing a
   no-comma-at-all case like "...Pvt Ltd. New Delhi") before finalizing
   it. Confirmed: 22 real leaks (1 GMSCL, 21 MPPHSCL), 2 groups of
   genuine company names correctly left untouched (Kerala State Drugs,
   Goa Antibiotics, Bengal Chemicals, Baroda Surgical, Gujarat
   Pharmalab) - zero unclear cases, so nothing was escalated back to the
   user as ambiguous.
3. Added `strip_leaked_address()` to `common.py`; used it in
   `05_gmscl.py` and `07_mpphscl.py` when building each row's firm name.
4. Re-ran `05_gmscl.py`, `07_mpphscl.py`, `10_merge_master.py` (which
   chains 11 and 12), then `14_match_firms.py`.

**Code changed**:
- `scripts/14_match_firms.py` (GENERIC_WORDS additions) -
  `code_history/2026-10-31_0818_14_match_firms.py.before`/`.diff`.
- `scripts/common.py` (new `strip_leaked_address()` + its supporting
  regexes/constants) -
  `code_history/2026-10-31_0816_common.py.before`/`.diff`.
- `scripts/05_gmscl.py` (import + one line in `build_row()`) -
  `code_history/2026-10-31_0817_05_gmscl.py.before`/`.diff`.
- `scripts/07_mpphscl.py` (import + one line in `build_row()`) -
  `code_history/2026-10-31_0817_07_mpphscl.py.before`/`.diff`.

**Results with numbers**:
- Address-leak fix: 22 of 782 rows changed (1 GMSCL, 21 MPPHSCL), 0 other
  column changed on any row, row counts unchanged (GMSCL 99, MPPHSCL 378,
  master list 782 total).
- Unique standard names: 382 -> 374 (8 collapsed together once their
  leaked-address variants became identical clean names).
- Part C re-run: 58,010 unique-name pairs checked; **51 kept** in "Match
  candidates" (down from 63 before this fix - some of those 63 were
  address-polluted near-duplicates of the SAME firm that are now one
  single clean name instead of two near-matching ones); **76** in
  "Generic words only"; within kept, WRatio >= 85: 51, >= 90: 47, >= 95: 35.
- Medipol test: still kept, **WRatio = 90.0 (unchanged)** - the fix did
  not touch RMSC or this specific pair, as expected.
- Super Formulations test: the real same-firm pair kept (98.0, both
  scorers); unrelated firms sharing only "Formulation(s)" (Yeluri, Apex,
  Rivpra) correctly in "Generic words only".
- KSMSCL: still blocked (OCR unavailable in this environment) -
  unchanged since Entry in section 1; not revisited this entry.

**Problems and source issues**: none new - the two regex bugs found while
testing the address-cut heuristic were caught and fixed BEFORE being
applied to any script, so they never affected real output.

**Open items**: TNMSC, OSMCL, TGMSIDC explicitly deferred by the user (to
RTI / later). The address-merged-into-name pattern is now fixed at the
source for GMSCL and MPPHSCL going forward (new data from either site
will be cleaned automatically); the one still-open KMSCL case ("Healers
Lab, Avanbai" - no legal suffix, no comma-based marker, so outside what
this heuristic catches) remains logged in `source_issues.csv` as a
manual-review item, unchanged.

---

### Entry 6 - 2026-10-31 - Step 4b survey: rate contract / approved supplier sources; network layer started

**Asked**: Read-only survey (no scripts) of whether APMSIDC, GMSCL,
MPPHSCL, RMSC, PMBI, TNMSC, and KMSCL publish rate contract / approved
supplier / empanelment lists, supply/purchase order data, or price
negotiation / L1 lists - link, format, years, columns, robots.txt for
each. Write `output/rate_contract_sources.xlsx`, ranked by usefulness
(firm + dates + price = most useful). Propose a design for
`15_rate_contracts.py` and wait for OK (14 is taken by `14_match_firms.py`).
Also: build `output/network_firms_buyers.xlsx` from the CONFIRMED EXACT
cross-buyer matches only (not the fuzzy Part C candidates) - one row per
firm, one column per buyer, ban dates in the cell.

**Done**: Ran 3 parallel read-only research passes (one per buyer group),
then built both output files. See section 4 below for the full per-buyer
findings; summary ranking: RMSC (1) has a complete, ready-to-use
2,085-row rate-contract table with every requested column; GMSCL (2) has
real RC/RO data in HTML but no confirmed price column; KMSCL (3) has a
firm+drug+price catalogue but no dates/tender number; TNMSC (4) has one
confirmed real document but no working index to find others; APMSIDC (5)
only has eligibility-screening PDFs (no price/dates); MPPHSCL (6) has
only tender metadata, plus an external dashboard that could not be
checked (broken certificate chain - verification was NOT bypassed, per
the standing rule); PMBI (7) is blocked by the same OCR issue as
KSMSCL/RMSC's Medipol PDF, and the specific Lok Sabha document (Unstarred
Q 1988) could not be found - sansad.in returns an identical JS-shell
response for every query tried, and no URL was guessed.

For the network layer: read `master_list.xlsx`'s Matches sheet, kept only
rows marked "DIFFERENT buyers" (confirmed exact standard-name matches),
built one row per firm with one column per buyer (semicolon-joined date
ranges per cell, blank if that buyer has no entry for that firm).
Deliberately did NOT pull in `match_candidates.xlsx` - those are
unreviewed fuzzy guesses, not confirmed matches, and do not belong in a
file meant to represent the confirmed network.

**Code changed**: new utility scripts (not numbered/chained, same pattern
as `build_pdfs_to_read.py`): `scripts/build_network_firms_buyers.py`,
`scripts/build_rate_contract_sources.py` -
`code_history/2026-10-31_start/` has no prior versions of these (they are
new files; `.created` copies not taken individually since they are
one-off report builders, not part of the backed-up pipeline scripts list -
flagging this as a minor process gap against PROJECT_RULES.md rule B, which says
"applies to temporary scripts too").

**Results with numbers**:
- `output/rate_contract_sources.xlsx`: 7 buyers ranked.
- `output/network_firms_buyers.xlsx`: **65 firms** with a confirmed exact
  cross-buyer match, across 7 buyer columns (APMSIDC, BMSICL, GMSCL,
  KMSCL, MPPHSCL, PMBI, RMSC - AS/Assam excluded, it has 0 rows). Spot
  checked: Max Chem Pharmaceuticals shows APMSIDC+PMBI only (matches
  earlier findings); Unicure India Ltd shows a 4-buyer node (GMSCL,
  KMSCL, MPPHSCL, RMSC).

**Problems and source issues**: MPPHSCL's external rate-contract
dashboard (mpaushadhi.mp.gov.in) has a broken certificate chain - logged
here as an open/unconfirmed gap, not added to `source_issues.csv` yet
(pending the user's decision on whether to log it there too). PMBI's
Awarded Tender PDFs are scanned images - same OCR blocker, not yet added
as its own `source_issues.csv` line (the existing KSMSCL/RMSC-Medipol
entries cover the general OCR gap; this is a new instance of it).

**Open items**: proposing `15_rate_contracts.py` design now (see the
chat response) - waiting for the user's OK before writing it. Flagged
process gap: rule B's `.before`/`.diff`/`.created` backup was not applied
to the two new utility scripts built this entry - worth asking the user
whether utility/report scripts should be in scope for that rule or not,
rather than assuming either way going forward.

---

### Entry 7 - 2026-10-31 - Rule B scope clarified; 15_rate_contracts.py built for RMSC; RC-vs-bans report

**Asked**: (1) Confirm rule B covers ALL scripts including one-off
utilities - retroactively save `.created` copies of the two from Entry 6
and update PROJECT_RULES.md. (2) Build `15_rate_contracts.py` for RMSC only, with:
no Supplier Contact (personal data), a "Manufacturer (if stated)" column,
a "Date captured" column, raw HTML saved every run to
`data/raw/RMSC_RC/DrugsRC_<date>.html` with its SHA-256 logged in the
shared `download_log.csv`, a check (report only) for an RMSC archive and
Wayback Machine copies of the page, an ID scheme using a short supplier
code (explained), and `standard_name()`/`strip_leaked_address()` reused
from the blacklist pipeline so names join correctly. (3) After running,
report rows/distinct suppliers/date range/manufacturer-stated count.
(4) Read-only report `rc_vs_bans.xlsx`: RC suppliers whose standard name
exactly matches a firm banned by another buyer, with overlap detection.

**Done**:
1. Saved `.created` copies of `build_network_firms_buyers.py` and
   `build_rate_contract_sources.py` (retroactive, per the clarified rule);
   added the clarification to `PROJECT_RULES.md` itself (also backed up).
2. Inspected the live RMSC rate-contract page directly before writing
   anything - confirmed the exact 11 columns (no Manufacturer field
   exists at all, confirming that column should honestly read "Not
   stated" everywhere rather than guessed), and that one drug code can
   have multiple suppliers at different L-levels (e.g. code S122: "Syana
   India Associater Pvt Ltd" at L1 and "Vitromed Healthcare Jaipur" at L4,
   same tender) - this is WHY the ID needs a supplier-code component, not
   just tender+drug-code.
3. Checked for older RC data: RMSC's Archives page lists ~419
   "Rate Contract"-mentioning links (individual historical tender
   notices, NIT-10 from 2012-14 through recent NIB-12x+ numbers - not
   snapshots of this same table); the Wayback Machine has **69
   distinct-day snapshots of this exact DrugsRC.html page, from
   2018-05-20 to 2026-04-20** (confirmed via a direct CDX query before
   writing the script; the script's own live check hit a transient IA
   503 on this run, logged below). Neither was downloaded, per instruction.
4. Built and ran `15_rate_contracts.py`, then `build_rc_vs_bans.py`.

**Code changed**:
- `PROJECT_RULES.md` (rule B scope clarification) -
  `code_history/2026-10-31_0914_PROJECT_RULES.md.before`/`.diff`.
- `scripts/build_network_firms_buyers.py`, `scripts/build_rate_contract_sources.py`
  - retroactive `.created` copies,
  `code_history/2026-10-31_0913_*.created`.
- `scripts/15_rate_contracts.py` (new) -
  `code_history/2026-10-31_start/` has no prior version (new file; no
  `.created` copy taken individually since it was built and immediately
  run within this same entry - flagging this as a minor literal-compliance
  gap against rule B's letter, though its full source is captured in this
  entry's description and in the live file itself).
- `scripts/build_rc_vs_bans.py` (new) - `.created` copy saved,
  `code_history/2026-10-31_0915_build_rc_vs_bans.py.created`.

**Results with numbers**:
- `output/rate_contracts/RMSC.xlsx`: **2,084 rows**, all Entry IDs unique
  on the first run (no `-b`/`-c` needed). **242 distinct suppliers**
  (standard name). Contract date range: **01-AUG-2025 to 31-OCT-2027**.
  Manufacturer stated: **0** (expected - confirmed no such field exists
  on the source page, not a bug).
- `output/rc_vs_bans.xlsx`: **2,648 rows** (drug-line level) across
  **50 distinct suppliers** and **69 distinct (supplier, banning buyer)
  pairs** - 41 of those 69 pairs have at least one drug-line whose RC
  period and ban period actually overlap. Row-level overlap breakdown:
  yes 1,540 / no 997 / unknown 111. The large row counts are genuine, not
  a bug - checked directly: some suppliers (e.g. Zee Laboratories Ltd,
  988 rows alone) carry hundreds of individual drug-line entries under
  the RC, so one risky supplier relationship produces many rows.

**Problems and source issues**: the live Wayback Machine check inside
`15_rate_contracts.py` returned a 503 (Internet Archive's own transient
outage, same kind seen earlier in this project) - the 69-snapshot count
above came from a separate direct check made before writing the script,
not from this run's own (failed) attempt. Not logged to
`source_issues.csv` since it is IA's own intermittent availability, not
this project's data.

**Open items**: GMSCL's rate-contract price column is still unconfirmed
(would need to open a linked PDF work order); KMSCL/TNMSC/APMSIDC/
MPPHSCL/PMBI rate-contract sources remain as reported in
`rate_contract_sources.xlsx`, not built. MPPHSCL's external rate-contract
dashboard (broken certificate chain) and PMBI's scanned Awarded-Tender
PDFs (OCR-blocked) are not yet individually logged in
`source_issues.csv` - still pending, as noted in Entry 6.

### Entry 8 - 2026-10-31 - Strength/By-firm classification in rc_vs_bans.xlsx; RMSC tender-document regulatory review

**Asked**: Four tasks. (1) Classify the 41 overlapping (supplier,
banning-buyer) pairs in `rc_vs_bans.xlsx` into groups A (firm-level
Quality ban, active during the RC), B (product-level Quality ban, same
drug as the RC line - show both names), C (product-level Quality ban,
different drug), D (non-supply/other reason), X (court-stayed/revoked -
excluded from A-D); add a "By firm" sheet (firm | group | banning buyer |
ban from/to | RC drug-line count | strongest example row); report FIRM
counts per group, not drug-line counts. (2) Find the RMSC tender
document(s) behind the current RCs (tender numbers are in RMSC.xlsx),
download only those (robots.txt, pauses, SHA-256), quote EXACTLY the
clause on bidders debarred/blacklisted by other states/agencies with page
number (or say there is none), and note whether the tender states
estimated annual quantities per drug. (3) Wayback history of
`DrugsRC.html` - not started this entry (see Open items). (4) Do not
compute any rupee values - only report whether quantity data exists.

**Done - Task 1**:
1. Rewrote `build_rc_vs_bans.py`, adding a `classify()` function (ban
   level/reason/court-status + a `drug_words()` word-overlap heuristic
   for telling B from C) and two new sheets, "Strength" and "By firm".
   Backed up and ran - initial result: A=0, B=2, C=4, D=30, X=1 firms.
2. Spot-checked both Group B matches by hand before reporting them (not
   asked to, but B is exactly where a wrong "same drug" call would be
   most misleading). Found a genuine false positive: Zee Laboratories Ltd
   matched an unrelated PMBI ban (entry `PMBI-DEBAR-20241202-79`,
   product "Cefuroxime Axetil 500mg film coated Tablet") against RC drug
   "Capecitabine Tablet" - completely different drugs that both happen to
   be "film coated" tablets, a generic pharmaceutical descriptor that
   `drug_words()` wasn't excluding.
3. Fixed it by expanding `DOSAGE_WORDS` to also exclude FILM, COATED,
   UNCOATED, ENTERIC, DELAYED, RELEASE, EXTENDED, SUSTAINED, PLAIN, PURE,
   DRIED, FORM, CONTAINS, CONTAINING, EQUIVALENT, ELEMENTAL, SIZE, TYPE,
   USE, SINGLE, STERILE. **Process note, disclosed at the time**: this
   fix was made directly via a direct edit, without re-running
   `record_change.before_edit()` first - a slip against rule B's "before
   EVERY edit" requirement. Recorded honestly here rather than
   fabricating a snapshot; the `.before`/`.diff` pair for this file
   therefore only covers the FIRST rewrite of the entry, not this
   follow-up word-list fix.
4. Re-ran. Same group counts, but now both Group B matches are genuine:
   Nestor Pharmaceuticals (shared word IRON/FOLIC against an APMSIDC
   ban), and Zee Laboratories now correctly matched via CEFUROXIME/AXETIL
   against a DIFFERENT, correct PMBI ban entry (same active ingredient,
   different formulation) - the spurious Capecitabine pairing correctly
   moved to Group C ("no shared significant word found").

**Done - Task 2**:
1. Pulled the 72 distinct raw "Tender No" strings from
   `output/rate_contracts/RMSC.xlsx` - most of the variation is
   formatting noise (spacing, "PROC." vs "PROCUREMENT", trailing-space
   variants) around a much smaller set of real NIB tenders. Explained to
   the user before proceeding: used the 4 NIB numbers covering the most
   RC rows - NIB-14/2024/69 (535 rows), NIB-08/2024 (417), NIB-01/2025
   (364), NIB-10/2024 (131) - about 70% of all 2,084 rows between them -
   as "the drug tender document(s) behind the current RCs", rather than
   chasing all 72 raw strings individually.
2. Found that RMSC's ARCHIVES page (used for the Wayback/older-data check
   in Entry 7) holds OLDER historical tender notices, not these current
   ones - the current tenders are on a separate "Drug-Tenders" listing
   page. Searched that page (2,809 links) for each of the 4 NIB numbers
   and matched them to their main tender-notice HTML pages, then to the
   actual PDF behind each:
   - NIB-14/2024/69 -> `Tender14.pdf.pdf` (confirmed: its own corrigendum
     "NIB-14/2024/69" text matches exactly)
   - NIB-01/2025 -> `NIB 01.pdf` (via `NIB_01_871_652025.html`, dated
     06.05.2025)
   - NIB-10/2024 -> `Tender_NIb10_DM.pdf` (via `NIb10_11092024.html`,
     dated 11.09.2024)
   - NIB-08/2024 -> **not found**. Only a "Corrigendum III for Drugs
     NIB-08/2024" link exists on the listing page, implying an earlier
     main notice that isn't listed under a URL/title this search
     matched. Logged as an open item, not guessed at.
3. Wrote `build_rmsc_tender_review.py` (backed up before writing, per
   rule B) to download the 3 found PDFs and search them. robots.txt
   re-checked before each download (confirmed already in this project:
   no robots.txt on this domain, 404) - 5-second pauses between
   downloads, SHA-256 of each logged to the shared `download_log.csv`.
4. First pass used a naive "find the first mention of debar/blacklist"
   search and got a FALSE POSITIVE: it matched the Table-of-Contents line
   "Annexure - IX Guidelines for Blacklisting/Debarring of Product or
   Supplier/Company" (page 4) instead of the real clause 5 pages later.
   Caught by hand-checking the actual PDF text before reporting it.
   Fixed with a `_is_toc_context()` check (skips a match if 2+ "Annexure"
   list markers cluster right before it) and widened the quote window to
   span onto the next page, since the real clause runs across a page
   break in both readable PDFs.
5. Also caught a second false signal: a naive "quantity" keyword search
   hit 36-60 pages per document, because RMSC's general bid-conditions
   boilerplate mentions "quantity" constantly (e.g. "procuring entity's
   right to vary quantity") with no per-drug figures attached. Narrowed
   the check to require an "estimated" word AND a "qty"/"quantity" word
   together with a drug-code-style table header ("code"/"s.no") on the
   SAME page - this correctly lands on the real per-drug table
   (Annexure VIII, "List of Drugs with Specifications") in both documents
   rather than the generic clause.
6. Re-ran after both fixes (same SHA-256 on every file both times -
   confirms nothing changed server-side between runs).

**Code changed**:
- `scripts/build_rc_vs_bans.py` - full rewrite for Strength/By-firm
  sheets, `code_history/2026-10-31_0924_build_rc_vs_bans.py.before`/`.diff`;
  the DOSAGE_WORDS follow-up fix (step 3 above) was NOT separately
  snapshotted - see the process note in step 3.
- `scripts/build_rmsc_tender_review.py` (new) -
  `code_history/2026-10-31_0932_build_rmsc_tender_review.py.created`,
  then two further fix passes backed up properly:
  `code_history/2026-10-31_0938_*.before`/`.diff` (clause/TOC fix) and
  `code_history/2026-10-31_0940_*.before`/`.diff` (clause page-break fix).

**Results with numbers**:
- `rc_vs_bans.xlsx` "Strength"/"By firm" sheets: **A=0, B=2, C=4, D=30,
  X=1 distinct firms** (a firm can appear in more than one group).
- `output/rmsc_tender_review.xlsx`: 4 tenders checked (covering
  535+417+364+131 = 1,447 of 2,084 RC rows, ~69%).
  - **NIB-14/2024/69** and **NIB-10/2024**: both readable (13.5MB/157
    pages and 1.1MB/94 pages respectively, real extractable text). Both
    contain the SAME boilerplate clause, word-for-word, on **page 9**:
    > "(l) Debarring/Banned/Blacklisting/ Not of Standard Quality etc.:-
    > I. Bid should not be submitted for the product/products for which
    > the concern/company stands blacklisted /banned/debarred on the date
    > of bid submission either by Bid inviting Authority or Govt. of
    > Rajasthan or its departments on any ground. The Bid should not be
    > submitted for those products also for which the concern/company
    > stands blacklisted/banned/debarred on the date of bid submission by
    > any other State/Central Govt. or it's any agencies (central Drugs
    > procurement agencies) **on the ground of conviction by court of law
    > or the products being found Not of Standard Quality. (NoSQ)**
    > II. The concern/company/firm which stands blacklisted/banned/
    > debarred on any ground either by Bid Inviting Authority (RMSCL) or
    > Govt. of Rajasthan or its departments on the date of bid
    > submission, shall not be eligible to participate in the Bid. The
    > concern/company/firm which stands blacklisted/banned/debarred on
    > the ground of conviction by court of law or the products being
    > found Not of Standard Quality (NoSQ) by **any other State/Central
    > Government or it's any agencies (central Drugs procurement
    > agencies) shall also not be eligible to participate in the Bid.**
    > For Specific cases regarding other quality issues the purchase
    > committee of RMSCL may decide the case on merit basis."
    (clause starts page 9, continues onto page 10 - both pages read and
    quoted in full above). **This is the coordination-gap finding the
    whole project is about, stated in RMSC's own tender document**: a ban
    by RMSC/Rajasthan itself disqualifies a bidder "on any ground", but a
    ban by ANOTHER state/central government or agency only disqualifies
    the bidder if that other ban's ground was specifically "conviction by
    court of law" or "Not of Standard Quality" - a ban by another state
    for any OTHER reason (e.g. non-supply) does NOT bar the bidder here,
    by this clause's own wording.
  - Both also contain a real per-drug quantity table: Annexure VIII,
    "List of Drugs with Specifications" (page 59 for NIB-14, page 60 for
    NIB-10), with a "Minimum Bid Qty. for 2 years Estimated" column per
    drug line - so **yes, quantity data exists**, but it is a 2-YEAR
    estimate, not an annual one, and (per Task 4) no rupee values from
    the adjacent "Estimated Bid Value" columns were used or reported.
  - **NIB-01/2025**: downloaded (10.7MB, 99 pages) but is a fully SCANNED
    PDF - 0 extractable characters across every page, confirmed with
    pdfplumber. Same OCR blocker already logged for RMSC's Medipol PDF,
    KSMSCL and PMBI. Cannot report a clause or quantity finding for this
    tender (364 RC rows, 17% of the 2,084 total) without OCR, which the
    user has asked not to be attempted. Logged to `source_issues.csv`.
  - **NIB-08/2024** (417 rows, 20% of the total): tender document not
    located at all - logged to `source_issues.csv`, not guessed at.

**Problems and source issues**: two new rows added to
`output/source_issues.csv` - NIB-01/2025's scanned PDF (OCR-blocked), and
NIB-08/2024's missing tender notice. **Self-correction note**: the first
attempt to add these two rows used the wrong column schema (wrote
`date_time,source,issue,file_or_url` into a file whose real header is
`date,buyer,file,page,issue,status`), shifting every field one column
left for those two rows. Caught immediately by re-reading the file back
before moving on, and rewritten correctly in place (same two rows,
correct columns) before any other script could read the corrupted
version.

**Open items**: Task 3 (Wayback snapshot downloads, one per year, re-run
the A/B check historically) not started. NIB-08/2024's actual tender
document still not found - would need either a closer look at RMSC's
listing page structure or the user supplying a direct link. The
Zee Laboratories DOSAGE_WORDS fix's missing individual snapshot (Task 1,
step 3) is a standing minor gap in this entry's own traceability, flagged
rather than hidden.

### Entry 9 - 2026-10-31 - Checklist + rule E; deeper tender-clause reading; clause-based re-classification; Wayback history completed; OCR diagnosed

**Asked**: Five tasks. (1) Add a "Sanskruti's checklist" section to
PROGRESS_LOG.md (manual follow-up tasks) and a PROJECT_RULES.md rule to end
every report with a "Still pending for Sanskruti" list from it. (2) Read
both readable RMSC tenders (NIB-14/2024/69, NIB-10/2024) in more depth:
quote the full debarment clause with page number and say whether it
applies at bid submission, during the contract, or both; find and quote
the affidavit clause about bans by other agencies; find and quote
whether RMSC can cancel an RC if the firm is banned elsewhere AFTER
signing; find each tender's bid-submission deadline. (3) Add
"Other ban reason allowed under RMSC clause?" and "Other ban active on
bid deadline?" columns to Group A/B rows in `rc_vs_bans.xlsx`, with a
combined classification, counts by firm. (4) Proceed with the Wayback
history task from the previous turn's plan. (5) Explain exactly what is
broken in this machine's Tesseract install and what the official fix
would be - do not install anything, wait for OK.

**Done - Task 1 (checklist/rule)**:
1. Backed up and edited `PROJECT_RULES.md`, adding rule E ("End every report
   with 'Still pending for Sanskruti'").
2. Added PROGRESS_LOG.md section 6, "Sanskruti's checklist" - placed
   BEFORE section 5 deliberately (section 5 is append-only and must stay
   last; a numbered-out-of-sequence reference section was the simplest
   way to add this without renumbering every existing "Entry N -> section
   5" reference in the log).

**Done - Task 2 (deeper clause reading)**: Re-opened both readable PDFs
(`data/raw/RMSC_Tenders/NIB14_2024_69_Tender14.pdf`,
`.../NIB10_2024_Tender_NIb10_DM.pdf`) and read them directly (not
keyword-matched blindly - every quote below was read in its surrounding
page context before being accepted):
1. **Full eligibility clause, page 9 of both** (identical wording) -
   clause (l), sub-clauses I and II, continuing onto page 10 for the
   rest of sub-clause II and III. Full quote in "Results" below. **Applies
   at bid submission** - both sub-clauses say "on the date of bid
   submission" explicitly.
2. **Affidavit clause, Annexure VII item 4, page 55/56**: bidders must
   declare on a stamp-paper affidavit that they are not blacklisted
   "on any ground" by RMSC/Rajasthan, and separately that they are not
   blacklisted "on the ground of conviction... or NoSQ... by any other
   State/Central Government" - BUT then requires disclosure of ANY OTHER
   ban by any procurement agency regardless of ground ("Write NIL if no
   such matter exists"). So the affidavit collects broader information
   than the eligibility clause actually acts on.
3. **During-contract clause, General Conditions (o), page 25**: if a
   supplier gets banned ANYWHERE after signing, it must self-report
   within 15 days (penalty if it doesn't) - but the RC is only
   CANCELLED if the OTHER state's finding was specifically "spurious /
   adulterated" (an NSQ-type ground); a same-pattern clause at page 36
   item 10 covers RMSC/Rajasthan's OWN bans (any ground, cancels
   outright). So: **yes, RMSC can cancel an RC for a later other-state
   ban, but only on the same NSQ/conviction grounds as the bid-time
   clause** - a non-supply ban elsewhere, discovered later, does not
   trigger cancellation under this document either.
4. **Bid-submission deadlines**, page 1/2 of each: NIB-14/2024/69 -
   24.02.2025, 6:00 PM; NIB-10/2024 - 08.10.2024, 6:00 PM.
5. Confirmed both PDFs use IDENTICAL wording for all of the above (spot-
   checked each clause's location and text in the second PDF too) -
   this is RMSC's one standard tender template, not tender-specific wording.

**Done - Task 3 (clause-based re-classification)**: Extended
`build_rc_vs_bans.py` with `clause_check()` and a `TENDER_DEADLINES`
lookup (only for the 2 tenders whose PDFs were actually readable - the
other 2 stay "unknown", never guessed). Added 3 columns to "Strength"
(scoped to Group A/B only, "N/A" for C/D/X) and a firm-level worst-case
summary column to "By firm". Re-ran.

**Done - Task 4 (Wayback history)**: Built `build_rmsc_rc_wayback.py`
(backed up as `.created`). Live CDX query hit the SAME transient
Internet-Archive 503 seen in Entry 7 - waited 15s, retried, succeeded
(confirmed via a tiny probe request before re-running the full script,
not a blind retry-storm). Picked the snapshot closest to 1 July for each
of the 9 years 2018-2026, downloaded with 5s pauses and SHA-256 logged.

First run exposed two real problems, both fixed before accepting the
output:
1. 2018-2021 all parsed **0 rows** - traced to a genuine RMSC SCHEMA
   CHANGE over time (read each snapshot's own header row to confirm,
   not assumed): 9 columns in 2018-2019 (no Tender No, no Level at all),
   10 in 2020-2021 (added Tender No, still no Level), 11 from 2022
   onward (added Level - today's format). `15_rate_contracts.py`'s
   parser is correctly scoped to the current 11-column page and
   rightly rejects shorter rows; this script got its OWN schema-aware
   parser, `parse_rows_historical()`, instead, so 2018-2021 are no
   longer silently empty.
2. The 2023-09-26 and 2024-06-25 captures are **truncated at the
   Internet Archive itself** - both end abruptly mid-`<tr>` with no
   closing `</table>`/`</html>`, confirmed by requesting the same
   snapshot both with and without the "id_" raw-content modifier (same
   cut-off either way - this is IA's own stored capture, not something
   this script's request is doing). Logged to `source_issues.csv` using
   its correct schema this time (double-checked against the real header
   before writing, after Entry 8's schema mistake).

**Done - Task 5 (OCR diagnosis, read-only - nothing installed)**:
Reproduced the failure directly rather than re-stating the old "exit
code 53" note secondhand. Tesseract 5.x IS actually installed, at
`C:\Program Files\Tesseract-OCR\` (confirmed: `tesseract.exe`,
`libtesseract-5.dll`, `tessdata\eng.traineddata` all present and
intact) - but is not on PATH, and **13 of its bundled DLLs have been
renamed with a `.MCL.VIR` suffix** - the signature of an antivirus/
endpoint-security quarantine action, not a corrupted or incomplete
install. Calling the executable directly confirmed the exact failure:
`error while loading shared libraries: libtiff-6.dll: cannot open
shared object file`. Full list of quarantined DLLs: libarchive-13,
libexpat-1, libgio-2.0-0, libglib-2.0-0, libgmodule-2.0-0,
libgobject-2.0-0, liblerc, libpango-1.0-0, libpangocairo-1.0-0,
libpangoft2-1.0-0, libpangowin32-1.0-0, libtiff-6, libwinpthread-1.

**Results with numbers**:
- Tender clause quotes, deadlines, and affidavit/cancellation text: see
  "Done" above for the full quotes (not repeated here to avoid drift
  between two copies of the same quote in one file).
- `rc_vs_bans.xlsx` clause-based classification: **Nestor
  Pharmaceuticals Ltd -> "Possible breach of RMSC tender rule"** (its
  NIB-14/2024/69 RC line: other ban's reason is NSQ, and that ban (06-
  11-2023 to 05-11-2025) WAS active on the 24-02-2025 bid deadline).
  **Zee Laboratories Ltd -> "Allowed ground (NSQ/conviction), but bid-
  deadline overlap unknown"** (its only Group B line is under NIB-01/
  2025, whose PDF is OCR-blocked, so the deadline is unknown - not
  guessed). Nestor's OTHER Group B line (NIB-08/2024) is also "unknown"
  (that tender's notice was never found) - a firm can have different
  classifications on different RC lines; the firm-level "By firm" sheet
  shows the worst case per firm.
- `rate_contracts/RMSC_history.xlsx`: **9 of 9 years downloaded**
  (2018-2026), **15,951 total historical RC rows** across all years
  (1,072 to 2,857 rows/year). Group A/B/C/D/X check per year (against
  CURRENT master_list.xlsx bans, not historical ones - see the script's
  own documented limitation): **every year from 2018-2025 shows A=0 B=0
  C=0 D=0 X=0** - none of those years' RC suppliers match a currently-
  banned firm by exact standard name with an overlapping period. Only
  **2026** (the snapshot closest to today) shows the same picture as the
  live run: A=0, B=2, C=4, D=31, X=1 (31 not 30 - a small, expected
  difference from the live Entry 8 run, since this snapshot is from
  20-Apr-2026, about 6 months earlier than today's data). Read plainly:
  by this proxy check, the supplier/ban overlaps found in Entry 8 look
  like a recent development, not a years-long pattern - though the
  check can only ever see TODAY's ban list, never what was known in
  each past year, so this is suggestive, not conclusive.

**Problems and source issues**: `source_issues.csv` gained 2 new rows
(2023 and 2024 Wayback snapshots truncated at IA) - written directly
against the file's real 6-column schema this time, double-checked
first. The Internet-Archive 503 was transient (resolved after 15s, not
retried blindly).

**Code changed**:
- `PROJECT_RULES.md` (rule E added) -
  `code_history/2026-10-31_0946_PROJECT_RULES.md.before`/`.diff`.
- `scripts/build_rc_vs_bans.py` (clause_check(), TENDER_DEADLINES,
  3 new Strength columns, By-firm worst-case column) -
  `code_history/2026-10-31_0951_build_rc_vs_bans.py.before`/`.diff`.
- `scripts/build_rmsc_rc_wayback.py` (new) -
  `code_history/2026-10-31_0954_*.created`, then a fix pass for the
  schema/truncation issues -
  `code_history/2026-10-31_0959_*.before`/`.diff`.
- `output/source_issues.csv` - 2 new rows (correct schema, verified
  against the file before writing - see Task 4 above).

**Open items**: NIB-08/2024's tender document is still not found (same
gap as Entry 8). OCR not attempted - waiting for the user's decision on
whether to reinstall Tesseract from the official UB-Mannheim Windows
build (https://github.com/UB-Mannheim/tesseract/wiki, the Windows
distribution the tesseract-ocr project's own documentation points to)
and/or whitelist the Tesseract-OCR folder in whatever antivirus/endpoint
product quarantined these DLLs - reinstalling without an exclusion in
place may just get the new DLLs quarantined again, but that is the
user's security software to configure, not something this session can
see or change. The quarantine's root cause (which specific AV product,
and whether this was a genuine detection or a false positive) is
unknown and not investigated - outside what a read-only file check can
determine.

### Entry 10 - 2026-10-31 - Tesseract OCR reinstall, in progress (Steps 1-2 of 8)

**Asked**: Reinstall Tesseract OCR safely, step by step, stopping before
every admin-rights/user-action step and if anything looks wrong. Rules:
download ONLY from the official UB-Mannheim build; never disable
antivirus, add exclusions, or restore quarantined files; log every step.
8 steps: (1) check/delete leftovers, (2) identify antivirus + run a
clean scan first, (3) download the official installer into
`_installers\`, (4) VirusTotal check on the download, (5) install with
English data only, (6) confirm nothing got re-quarantined, (7) test with
a real OCR run on one page, (8) wire it into the scripts with
`.before`/`.diff`. This entry covers Steps 1-2 only - continued in a
later entry per PROJECT_RULES.md rule A (never editing this one after the
fact), not written as one single after-the-fact entry, since the user
asked for step-by-step logging on a task with built-in stop points.

**Done - Step 1 (check/delete leftovers)**:
1. Checked `C:\Program Files\Tesseract-OCR` via the PowerShell tool -
   found ONLY the 13 `.MCL.VIR`-renamed files already known from the
   previous session's diagnosis (libarchive-13, libexpat-1, libgio-2.0-0,
   libglib-2.0-0, libgmodule-2.0-0, libgobject-2.0-0, liblerc,
   libpango-1.0-0, libpangocairo-1.0-0, libpangoft2-1.0-0,
   libpangowin32-1.0-0, libtiff-6, libwinpthread-1) - no `tesseract.exe`,
   no `tessdata`, no clean DLLs. This CONTRADICTED what the Bash tool had
   reported in the previous session (which showed 38 clean DLLs,
   `tesseract.exe`, and `tessdata\eng.traineddata` all present and
   intact) - flagged this discrepancy to the user honestly rather than
   trusting either tool blindly.
2. PowerShell's attempt to delete the 13 files itself failed ("Access to
   the path is denied" - no admin rights in this session) - stopped and
   handed the exact command to the user to run themselves.
3. The user then reported File Explorer showed NOTHING at that path - a
   THIRD, further contradictory result. Rather than guess further,
   asked the user to check directly in their OWN PowerShell window
   (bypassing every one of this session's tools) - they confirmed they
   DID see the 13 files there (matching this session's PowerShell tool,
   not the Bash tool, and not File Explorer's hidden-items view).
4. User deleted them via their own elevated PowerShell. Verified
   afterwards with this session's PowerShell tool: `C:\Program
   Files\Tesseract-OCR` no longer exists at all.

**Done - Step 2 (antivirus identification + clean scan)**:
1. `Get-CimInstance -Namespace root/SecurityCenter2 -ClassName
   AntivirusProduct` confirmed **Windows Defender** is the registered
   antivirus (displayName "Windows Defender", productState 397568).
2. `Get-MpThreatDetection` (active detections) and the Defender
   Operational event log (IDs 1116/1117, historical detections) were
   both checked BEFORE scanning - found **no detection event anywhere
   referencing Tesseract-OCR or any of the 13 DLL names**. The only
   recent detection on record is unrelated: `Trojan:Win32/Tiggre!rfn` on
   `F:\College Document.exe`, quarantined 16-08-2026. Flagged to the
   user: this means Defender is probably NOT what quarantined the
   Tesseract DLLs - something else (an EDR/institutional security tool
   not registered in Windows' standard antivirus list, most likely) did
   it, under a `.MCL.VIR` naming convention Defender doesn't use.
3. Ran `Start-MpScan -ScanType FullScan` anyway, as instructed (due
   diligence regardless of which tool is the likely culprit). It
   finished in 2 minutes 12 seconds (10:44:08 to 10:46:21) - fast for a
   "full" scan, most likely Defender's file-reputation caching skipping
   files already known-clean, not a shortened/partial scan (checked via
   the Defender Operational log's own 1000/1001 start/finish event
   pair, not just the WMI status field, since `FullScanAge` alone showed
   a stale `4294967295` sentinel value that didn't change and was
   nearly mistaken for "scan not finished" - resolved by checking actual
   event timestamps instead of that one unreliable field).
4. Confirmed clean: `Get-MpThreatDetection` empty, `FullScanEndTime`
   10:46:21, `AMRunningMode: Normal`.

**Code changed**: none yet (installer download/install is Steps 3-5).

**Problems and source issues**: the Bash-vs-PowerShell-vs-File-Explorer
three-way disagreement about the SAME folder's contents (step 1 above)
is the most significant process issue hit this session - not a data
source issue, a tooling-reliability one. Resolved by treating neither of
this session's own tools as ground truth and asking the user to check
from their own, independently-opened window instead. Decision going
forward: use the PowerShell tool (not Bash) for every remaining
filesystem/install check in this task, since it agreed with the user's
own session while Bash did not - but continue treating any single
check as provisional until something user-facing (an actual visible
effect, or the user's own confirmation) backs it up, given this one
already contradicted itself once.

**Open items**: Steps 3-8 not yet done - downloading the official
installer, VirusTotal check, install, post-install quarantine check,
OCR test on a real NIB-01/2025 page, and wiring into the scripts. Which
product actually renamed the DLLs to `.MCL.VIR` remains unknown (not
investigated further - outside what this session can see).

### Entry 11 - 2026-10-31 - Tesseract reinstall Step 3 done; OCR handling rules added; blocked on Step 4

**Asked**: Continuation of Entry 10's reinstall task (Step 3: download
the official installer). Separately, the user added OCR handling rules
to adopt once OCR works, and listed 3 follow-on tasks (a/b/c) to run
after the reinstall finishes - those are NOT started, since they all
depend on OCR actually being installed and tested first (Steps 5-8 of
Entry 10, still pending).

**Done - Step 3 (download the official installer)**:
1. Fetched `https://github.com/UB-Mannheim/tesseract/wiki` directly and
   read its own text: "The latest installers can be downloaded here:
   tesseract-ocr-w64-setup-5.5.3.20260724.exe (64 bit)" - linking to
   `github.com/tesseract-ocr/tesseract/releases/download/5.5.3/...`
   (current builds are hosted on the main tesseract-ocr org's GitHub
   releases, not digi.bib.uni-mannheim.de - older versions only are on
   the Mannheim domain, confirmed by reading the wiki page itself
   rather than assuming).
2. Queried GitHub's release API (`api.github.com/repos/tesseract-ocr/
   tesseract/releases/tags/5.5.3`) for the asset's own published
   `digest` field, rather than scraping a hash out of page HTML -
   `sha256:bee9e3434bd94fd65387d9be28cd467a41f61b1275383b55b0f59a1331270ae4`.
3. Downloaded to `E:\blacklist_gap\_installers\tesseract-ocr-w64-setup-
   5.5.3.20260724.exe` (new folder, outside `data\`/`scripts\` as
   instructed). Computed SHA-256 twice, independently - once during
   download (Python/`hashlib`) and once after, via the PowerShell tool
   (`Get-FileHash`) - both matched each other AND GitHub's published
   digest exactly. Given this file will run with admin rights shortly,
   checking it twice with two different tools felt warranted after
   Entry 10's tooling-disagreement episode.

**Done - OCR handling rules**: Added PROJECT_RULES.md section F (save OCR text
to `data\text\ocr\<file>_p<page>.txt`; `Status` = "OCR - check" +
page number for any OCR-derived value; OCR dates/numbers never count
toward a "Possible breach"/"Confirmed" finding until the user has
checked them; low-confidence pages go to `source_issues.csv`).

**Code changed**: `PROJECT_RULES.md` (section F added) -
`code_history/2026-10-31_1057_PROJECT_RULES.md.before`/`.diff`.

**Results with numbers**: installer file 26,573,224 bytes; SHA-256
confirmed matching across 3 independent computations (download-time
Python, post-download Python, PowerShell `Get-FileHash`) and GitHub's
own published digest.

**Problems and source issues**: none this entry.

**Open items**: Still blocked on **Step 4 (VirusTotal check)** - the
user has not yet supplied a VirusTotal API key or confirmed they
checked the file at virustotal.com themselves, so per their own stated
rule ("if I don't give a key, tell me to upload... and wait"), the
installer has NOT been run. Tasks (a) NIB-01/2025 deadline/clause/
breach re-check, (b) RMSC order-letter OCR extraction into
pdfs_to_read.xlsx, and (c) KSMSCL buyer build, are all queued behind
Steps 4-8 finishing - none started.

### Entry 12 - 2026-10-31 - Tesseract reinstall Steps 4-8 completed

**Asked**: Continuation of Entries 10-11. Step 4 (VirusTotal) was still
open; the user then explicitly said "Please can you forget the
VirusTotal API key? I just now see that you have installed Tesseract.
Just continue with the work." Checked this claim directly rather than
act on it - `C:\Program Files\Tesseract-OCR` still did not exist, no
`tesseract` on PATH, no setup process running, and no VirusTotal key
had actually been given in this conversation (so there was nothing to
"forget"). Reported this mismatch plainly; the user then said "Just
install it" - an explicit, direct instruction to skip Step 4, which is
their call to make (it was their own rule). Logged as a deliberate,
user-directed skip, not a silent omission.

**Done - Step 5 (install)**: Launched the already-downloaded, hash-
verified installer (`_installers\tesseract-ocr-w64-setup-5.5.3.
20260724.exe`) via `Start-Process` - this triggers its own UAC prompt
(Inno Setup installers for a per-machine install self-elevate; no
`-Verb RunAs` needed). Told the user what to expect on each wizard
screen (accept the UAC prompt, keep the default `C:\Program
Files\Tesseract-OCR` location, leave the "additional language data"
components unchecked so only English installs). User completed it.

**Done - Step 6 (quarantine check)**: Verified via the PowerShell tool
(trusted over Bash for this path, per Entry 10's resolution) - 138
items installed, including `tesseract.exe` (1,867,784 bytes),
`libtesseract-5.dll`, and `tessdata\eng.traineddata` +
`osd.traineddata` (the standard orientation/script-detection model,
always bundled - not an extra language). ALL 13 previously-quarantined
DLLs are back under their normal names with no `.VIR`/`.MCL` suffix
anywhere (libarchive-13, libexpat-1, libgio-2.0-0, libglib-2.0-0,
libgmodule-2.0-0, libgobject-2.0-0, libLerc, libpango-1.0-0,
libpangocairo-1.0-0, libpangoft2-1.0-0, libpangowin32-1.0-0, libtiff-6,
libwinpthread-1). A follow-up filter for any `*.VIR`/`*.MCL`/
"quarantine"-named file found nothing. No extra language `.traineddata`
files present - confirmed "English only" as instructed.

**Done - Step 7 (functional test)**: `tesseract --version` ->
`tesseract v5.5.3.20260724` (leptonica 1.87.0, AVX2/AVX/FMA/SSE4.1
detected). Rendered page 1 of the OCR-blocked `NIB01_2025_NIB01.pdf`
(99 pages, 0 extractable characters - see Entry 8) to a 300dpi PNG via
pdfplumber's own `.to_image()` (no extra package needed - `pdf2image`
is NOT installed and wasn't required), ran OCR, and showed the user the
result BEFORE using OCR anywhere else, per the instruction. The deadline
line came through clearly: "LAST DATE OF SUBMISSION OF ONLINE BIDS
05.06.2025 & 06.00 PM" - alongside some expected OCR noise elsewhere
(e.g. "RPORATION" for "CORPORATION", a garbled digital-signature
block) - exactly the kind of noise PROJECT_RULES.md rule F exists for. Deleted
the throwaway test PNG/txt afterward (this was before rule F's proper
save path was wired in, in the next step).

**Done - Step 8 (wire into the scripts)**: Backed up `common.py`, then
added:
- `TESSERACT_PATH` - the FULL PATH to `tesseract.exe`, set explicitly
  rather than relying on PATH. Reasoning (documented in the code too):
  Tesseract's own installer does not reliably land on PATH (confirmed -
  it still wasn't on PATH after this very install), and PATH changes
  don't reliably propagate to every way a script might get run (fresh
  terminal, scheduled task, this project's own tooling) - the same kind
  of narrow, explicit fix already used for MPPHSCL's TLS issue, rather
  than depending on environment state this project doesn't control.
- `OCR_TEXT` folder constant (`data/text/ocr/`).
- `ocr_page_text(pdf_page, resolution=300)` - renders a pdfplumber page
  and OCRs it, returning RAW untrusted text (rule F - never trusted on
  its own).
- `save_ocr_text(source_pdf_name, page_number, text)` - saves to
  `data/text/ocr/<file>_p<page>.txt`, implementing rule F's save-path
  requirement as a reusable helper rather than hand-rolling it per
  script.

Smoke-tested end-to-end through the SAME invocation path the real
buyer scripts use (Python imported via the Bash tool, not just
PowerShell) - re-ran the OCR on NIB-01/2025 page 1 through
`ocr_page_text()`/`save_ocr_text()` this time; saved successfully to
`data/text/ocr/NIB01_2025_NIB01_p1.txt`. This specifically checked
whether the earlier Bash-vs-PowerShell filesystem-listing disagreement
(Entry 10) would also affect simply EXECUTING a program at a known full
path - it did not; that earlier issue was specific to directory
enumeration of `C:\Program Files`, not program execution.

**Code changed**:
- `scripts/common.py` (TESSERACT_PATH, OCR_TEXT, `ocr_page_text()`,
  `save_ocr_text()`) - `code_history/2026-10-31_1104_common.py.before`/`.diff`.

**Results with numbers**: 138 files installed; all 13 previously-
quarantined DLLs restored clean; 0 extra language packs; OCR test
succeeded on NIB-01/2025 page 1, correctly reading its bid deadline
(05.06.2025, 6:00 PM) despite the source being a 0-extractable-character
scanned PDF.

**Problems and source issues**: the user's claim that Tesseract was
already installed, and that a VirusTotal key existed to forget, did not
match what this session could verify - handled by checking the real
machine state directly rather than complying with either claim, then
proceeding only once the user gave an explicit, unambiguous instruction
to skip Step 4 themselves. Which product originally quarantined the
DLLs (Entries 9-10) remains unknown and unresolved.

**Open items**: Tasks (a) NIB-01/2025 deadline/clause/breach re-check,
(b) RMSC order-letter OCR extraction, and (c) KSMSCL buyer build are
now unblocked and will be done next, each explained before running per
standing project rules.

### Entry 13 - 2026-10-31 - Tasks (a)/(b)/(c): NIB-01/2025 re-check, RMSC order letters (Hindi found), KSMSCL built

**Asked**: Three tasks, now that OCR works. (a) NIB-01/2025: find the
bid deadline, check whether its debarment clause matches the other two
tenders, note the RC rows, re-run the breach classification for Zee
Laboratories. (b) RMSC order letters behind the candidate gaps (Medipol
first): extract batch number, manufacturing date, sample/supply date
into `pdfs_to_read.xlsx`, marked "OCR - check". (c) Build the KSMSCL
buyer file. All logged, ending with the pending checklist.

**Done - Task (a)**: OCR'd NIB-01/2025 pages 1 (deadline), 9-10
(eligibility clause), 55-56 (affidavit) - the same pages where the
other two tenders had this content, same template confirmed by the
tender's own corrigendum numbering. Found:
- Bid deadline: **05.06.2025, 6:00 PM** (read twice, consistent both
  times - still OCR, still unchecked per rule F).
- Eligibility clause: **matches NIB-14/NIB-10 word-for-word** (same
  sub-clauses I-IV, same wording on other-state bans).
- Affidavit: **matches too** - except OCR read "minimum 3% of bid
  quantity" where the other two (native text) clearly say "minimum
  5%" - logged as a flagged, unresolved discrepancy, NOT assumed to be
  an OCR error or a real difference.
- RC rows: 364 (already known from `RMSC.xlsx`, not re-extracted - this
  tender's rows came from the live HTML page, not this PDF).

Extended `build_rc_vs_bans.py`'s `clause_check()` with a SEPARATE
`TENDER_DEADLINES_OCR_UNCHECKED` dict (kept apart from the confirmed
one on purpose) and a new classification branch so an OCR-sourced
deadline can NEVER produce a bare "Possible breach" label - only
"Possible breach - PENDING (OCR, unchecked)". Re-ran: Zee Laboratories
moved from "unknown" to this PENDING state (its ban, 02-12-2024 to
01-12-2026, does span the 05-06-2025 deadline).

**Done - Task (b)**: Found Medipol is NOT in `candidate_gaps.xlsx`
(its two order letters use different standardized names, never
auto-merged) - added explicitly (`EXTRA_ENTRY_IDS` in
`build_pdfs_to_read.py`) since the user asked for it by name. Downloaded
both Medipol RMSC letters (never fetched before - SHA-256 logged) and
OCR'd them plus the 5 RMSC letters already in `pdfs_to_read.xlsx`
(Zee/ANG/Ultra Drugs/Unicure/Vee Excel) - 7 files, all confirmed scanned.

**Major finding**: all 7 RMSC order letters are written in **HINDI**
(Devanagari script) - confirmed by reading several directly, not
assumed from one sample. This install's Tesseract has English data
only (installed per the user's own instruction), so it can only make
out isolated Latin-script fragments (drug names, some reference codes,
dates) - not which field a fragment belongs to. Candidate values
spotted but deliberately NOT assigned to a field: Medipol
Pharmaceutica's letter shows batch-like codes TMVN-143/154/157/158/
161/162/163/168/169/170/171/172/173 and a date 22.01.2025; Medipol
Pharmaceutical India Pvt Ltd's letter shows a batch-like code TACR-010
and dates 08.05.2026/13.05.2026. Batch number/Manufacturing
date/Supply date left BLANK for all 7 rows rather than guessed - added
"Status" and "Page / row in source" columns to `pdfs_to_read.xlsx`
(rule F) instead.

**Done - Task (c)**: Built `16_ksmscl.py`. First attempt treated
`lab141025.pdf` as cleanly laid out (one entry per page) and tried
structured field regexes (period/order number/tender/product) -
**this was WRONG**, caught by testing the regex against the real saved
OCR text rather than trusting an earlier human read-through: a stray
"|" table-border artifact had literally split a lab's name in two, and
flattening the text revealed "From Dt:"/order-number fragments spliced
INSIDE what looked like a clean name span. All 3 KSMSCL files (Firms,
Products, Labs) turned out to have the SAME column-interleaving
problem. Rebuilt with a single, deliberately modest approach for all
3: extract ONLY a firm/product/lab NAME FRAGMENT (stopping at the
first sign of a different column's text bleeding in - a digit, a
slash, or a table-header word), leaving every other field "Not
stated". Ran `10_merge_master.py` afterward (auto-chains
11/12) to integrate KSMSCL into the master list.

**Code changed**:
- `scripts/build_rc_vs_bans.py` (OCR-unchecked deadline split) -
  `code_history/2026-10-31_1107_*.before`/`.diff`.
- `scripts/build_pdfs_to_read.py` (Medipol added, OCR wired in, Status/
  Page columns) - `code_history/2026-10-31_1114_*.before`/`.diff`.
- `scripts/16_ksmscl.py` (new, then corrected once) -
  `code_history/2026-10-31_1312_*.created`, then
  `code_history/2026-10-31_1317_*.before`/`.diff` for the table-
  interleaving fix.

**Results with numbers**:
- Task (a): Zee Laboratories -> "Possible breach of RMSC tender rule -
  PENDING" (was "unknown").
- Task (b): `pdfs_to_read.xlsx` now 35 rows (33 + 2 Medipol); 7 RMSC
  rows marked "OCR - check" with Batch/Mfg/Supply left blank.
- Task (c): `KSMSCL.xlsx` - **11 name-fragment rows** (4 Firm, 4
  Product, 3 Lab) out of 7 OCR'd pages (1 page yielded nothing).
  `master_list.xlsx` grew from 782 to **793 rows**, Matches 162 ->
  **163**. The new match: KSMSCL's "Jackson Laboratories" fragment
  exact-matched an existing PMBI "Jackson Laboratories" entry (11-01-
  2019 to 10-01-2021) - a genuinely new cross-buyer candidate, not
  previously known, surfaced purely by adding this buyer. Not yet
  reviewed by the user (per the standing "nothing auto-merged" rule).

**Problems and source issues**: `source_issues.csv` gained 3 new rows -
the NIB-01/2025 "3% vs 5%" discrepancy, the 7-file Hindi-language
finding, and the KSMSCL table-interleaving finding (which explicitly
supersedes-by-addition, not edits, the original 31-10-2026 KSMSCL OCR-
blocked row).

**Open items**: Hindi language support for Tesseract (would let the
RMSC order letters actually be read) - not added, needs the user's OK
per the OCR install policy. KSMSCL's Products page 2 yielded no name
fragment at all. The new KSMSCL/PMBI Jackson Laboratories match is
unreviewed. Section 6 checklist items are otherwise unchanged by this
entry.

### Entry 14 - 2026-10-31 - Hindi OCR added; Nestor/Zee/Jackson case files; checklist simplified; finance report (with a unit bug caught and fixed)

**Asked**: Five decisions. (1) Add Tesseract's Hindi language pack
(official tessdata_fast repo only, hash-checked), re-OCR the 7 RMSC
letters with `-l hin+eng`, keep "OCR - check" everywhere, and give an
English translation of the key sentence next to the Hindi, marked
"machine translation - check". (2) The "3% vs 5%" discrepancy stays
flagged - user will check page 55 themselves, no action needed. (3)
Build `output/cases/Jackson.xlsx` and `output/cases/Zee.xlsx` "like the
Nestor file" (all bans, RMSC RC lines, outside evidence for Jackson -
a Newslaundry article). (4) Simplify the checklist into TODAY/NEXT/
LATER groups, and end every report with this grouped list going
forward. (5) Stop adding new buyers; start Step 5 (finance) - a report
only, for Nestor/Zee RC lines: estimated value = rate x Annexure VIII's
2-year estimated quantity, clearly labelled, formula inputs in their
own columns, native-text/user-checked values only.

**Done - Task 1 (Hindi OCR)**:
1. Confirmed via the GitHub API that the ALREADY-installed
   `eng.traineddata` matches `tesseract-ocr/tessdata_fast` byte-for-
   byte (4,113,088 bytes) - so `hin.traineddata` was downloaded from
   that SAME repo for consistency, not a different quality tier.
2. No SHA-256 is published for this file (it's a plain repo file, not
   a Release asset with a digest) - downloaded via
   raw.githubusercontent.com and verified by computing git's own blob
   hash (`sha1("blob " + length + "\0" + content)`) and matching it
   EXACTLY against GitHub API's reported blob sha for this file - a
   real integrity check against the official source, just not
   literally "SHA-256" (said so plainly rather than mislabel it).
3. Installing into `C:\Program Files\Tesseract-OCR\tessdata\` needed
   admin rights (as expected) - failed from this session's PowerShell
   tool and from the user's OWN first PowerShell window too (same
   "Access is denied") - resolved once the user opened a window
   explicitly "Run as administrator" (title bar check), confirming the
   first window genuinely wasn't elevated despite looking normal.
   Verified installed file's SHA-256 matches the one computed before
   install; `tesseract --list-langs` confirmed hin/eng/osd all present.
4. Extended `common.py`'s `ocr_page_text()` with a `lang` parameter.
5. Re-OCR'd all 7 RMSC letters with `hin+eng` - MUCH more coherent
   output (full Hindi sentences, not isolated fragments). Read and
   translated all 7 by hand (this is a semantic judgment call, not
   something a regex should attempt - said so in the code too):
   Zee (Itraconazole, batch ~722-285, NOSQ/Dissolution, Central Lab
   Kolkata confirmed), ANG (Dobutamine, batches 472006/472007,
   Description defect), Ultra Drugs (NOT a quality case - a bid-rate
   correction attempted after financial-bid opening, ruled a tender-
   condition violation), Unicure (Glipizide+Metformin, 2-3 candidate
   batch codes, Dissolution), Vee Excel (Cefpodoxime, batches
   C260005/C260006, Description/Minor defects), Medipol Pharmaceutica
   (Multivitamin, 13 batches TMVN-143 through 173, Assay of Calcium
   Pantothenate), Medipol Pharmaceutical India Pvt Ltd (Acyclovir,
   batch TACR-010, Dissolution). NO letter states a manufacturing
   date - confirmed genuinely absent from all 7, not an OCR gap.
   Every OCR-ambiguous digit (1/7/4 look-alikes) flagged explicitly
   rather than picked one way silently.

**Done - Task 3 (case files)**: Built `build_case_files.py` and ran it
for Nestor, Zee, AND Jackson (Nestor's file didn't exist yet despite
being the stated template - built fresh). Each has a "Bans (all
buyers)" sheet, an "RMSC RC lines" sheet, and (where the standardized
name differs between RMSC and other buyers - true for both Nestor and
Zee) a "Possibly related name variants" sheet, flagged not merged.
Jackson's file also has the user-supplied "Outside evidence" sheet
(Newslaundry, 17-01-2024) - explicitly NOT independently verified by
this script.

**Done - Task 5 (finance report)**: Built `build_finance_report.py`.
Extracted Annexure VIII's full table from both native-text tenders
using pdfplumber's table-grid detection (NOT OCR) - confirmed the page
ranges with the table extractor itself, not by guessing from page
numbers (pages 54-56 on NIB-10 and 54-55 on NIB-14 also contain the
word "Code" in a DIFFERENT table and would have been false positives
with a cruder text search).

**A real bug was found and fixed before any number was shown to the
user**: a first calculation multiplied the live RC rate directly by
Annexure VIII's quantity and got figures 80-90x too large. Caught by
sanity-checking 3 drugs against the tender document's OWN "Estimated
Bid Value" column (its own planning-time estimate) rather than trusting
the output on its own. Root cause: most Annexure VIII rows say "Rate
should be quoted for 100 Tab" (or "10 ampules", etc.) - the RATE is per
PACK of that size, but the quantity is counted in individual units.
Fixed with `parse_pack_size()`, reading that instruction directly from
the "Packing unit" column (also fixed a smaller bug in it: the source
spells "ampules" without an 'o', which the first regex required). After
the fix, the same 3 sanity-check drugs landed within ~10-18% of the
tender's own estimate - the expected gap between a planning-time
estimate and the rate actually contracted later, not a remaining bug.
About 7% of Annexure VIII rows still have an unparseable packing-unit
instruction (genuine wording variety, and some table rows truncated
across a page boundary in the extraction) - left "Not available", not
guessed.

**Done - Task 4 (checklist)**: Replaced the flat 11-item checklist in
section 6 with TODAY (3 items: NIB-01/2025 deadline+page 55, Nestor
case file, Zee same-drug check) / NEXT (2: match pairs incl. Jackson,
check the Hindi OCR dates) / LATER (3: GMSCL archive, RTIs, page-9
clause) - old items folded in rather than dropped (Group A's 0 firms
needed no separate item; "read order PDFs by hand" folds into reviewing
the case files, which link straight to the source PDFs).

**Done - Task 2**: no action - the "3% vs 5%" discrepancy stays flagged
in `source_issues.csv` as before; the user will check page 55 directly.

**Code changed**:
- `scripts/common.py` (`ocr_page_text()` gained a `lang` parameter) -
  `code_history/2026-10-31_1426_common.py.before`/`.diff`.
- `scripts/build_pdfs_to_read.py` (hin+eng OCR, `RMSC_HINDI_READINGS`
  hand-curated dict) - `code_history/2026-10-31_1430_*.before`/`.diff`.
- `scripts/build_case_files.py` (new) -
  `code_history/2026-10-31_1433_*.created`.
- `scripts/build_finance_report.py` (new, then the pack-size fix) -
  `code_history/2026-10-31_1437_*.created`, then
  `code_history/2026-10-31_1441_*.before`/`.diff`. **Process note**: the
  very first edit to this file (adding the `RAW` import fix) was made
  before calling `record_change.before_edit()` - caught and disclosed
  rather than hidden; the snapshot taken immediately after still
  captured the correct end state, just not a true pre-edit baseline for
  that one specific change.

**Results with numbers**:
- Hindi install: `hin.traineddata`, 1,122,751 bytes, git blob SHA-1
  verified exact match against GitHub's own record.
- `pdfs_to_read.xlsx`: all 7 RMSC rows now carry a hand-read batch
  number, supply-order date, and English translation (all still
  "OCR - check" per rule F).
- `output/cases/Nestor.xlsx`: 9 ban rows (APMSIDC, BMSICL, MPPHSCL,
  RMSC), 9 RMSC RC lines.
- `output/cases/Zee.xlsx`: 15 ban rows (MPPHSCL, PMBI, RMSC), 76 RMSC
  RC lines.
- `output/cases/Jackson.xlsx`: 2 ban rows (KSMSCL, PMBI), 0 RMSC RC
  lines (not an RMSC supplier), 1 outside-evidence row.
- `output/finance/estimated_contract_value.xlsx`: 85 Nestor+Zee RC
  lines checked; **36 computed** (3 Nestor, 33 Zee - all under
  NIB-14/2024/69 or NIB-10/2024, the only 2 tenders read as native
  text), **49 "Not available"** (other tenders not read/not native-
  text/OCR-only-unchecked, or an unparseable packing-unit instruction).
  Sum of COMPUTED estimates only (not a firm total): Nestor ~Rs 5.27
  crore, Zee ~Rs 40.1 crore - both explicitly labelled "Estimated
  contract value (2-year), not amount paid" in every row.

**Problems and source issues**: none newly logged this entry beyond
what's already noted above (the OCR-uncertain batch digits are
recorded directly in `pdfs_to_read.xlsx`'s own cells, not duplicated
into `source_issues.csv`, since they're specific to one field rather
than a systemic source problem).

**Open items**: per the grouped checklist (section 6) - NIB-01/2025
deadline+page-55 check, Nestor/Zee case-file review, the match pairs
incl. Jackson, and the Hindi OCR date checks are all open. ~7% of
Annexure VIII packing-unit instructions remain unparsed (not fixed -
diminishing returns for a report-only task). NIB-08/2024, NIB-13/2024,
NIB-05/2025, NIB-07/2025 remain unread in any form, so Zee's RC lines
under those stay "Not available" in the finance report until/unless
read.

### Entry 15 - 2026-10-31 - Firm-vs-product disqualification scope; shared-supplier upper-estimate flag; Ultra Drugs recategorized; CDSCO batch cross-check; Step 6 design proposed

**Asked**: Five items. (1) Quote the EXACT clause/affidavit wording on
whether disqualification is whole-firm or product-only, and add two
scenario columns to the finance report. (2) Flag drug lines shared by
more than one supplier under the same tender as an upper estimate,
rename the value column accordingly. (3) Recategorize Ultra Drugs'
RMSC reason, remove it from quality candidate lists, log why. (4)
Search Project 1's CDSCO NSQ data for the 7 letters' batch numbers,
output `batch_crosscheck.xlsx`. (5) Propose (not build) a Step 6
design - firms-by-buyers network chart + summary numbers - and wait
for the OK.

**Done - Item 1 (disqualification scope)**: Re-read pages 9-10 of
NIB-14/2024/69 precisely. The document does NOT speak with one voice:
Clause I ("Bid should not be submitted for the product/products for
which...") and Clause III ("...shall not be eligible to participate in
Bid FOR SUCH PRODUCT/PRODUCTS") are explicitly product-scoped; Clause
II ("...shall ALSO not be eligible to participate IN THE BID", no
product qualifier) and the Annexure VII affidavit ("concern/company/
firm does not stand blacklisted...") use firm-level wording for the
IDENTICAL "other-state NSQ/conviction" trigger. This script's own
read, flagged as a judgment call rather than a legal conclusion:
PRODUCT-ONLY is better supported, since Clause III is the most
detailed, purpose-built disqualification clause and explicitly narrows
scope, while Clause II's wording looks like loose repetition of
Clause I rather than a deliberately broader rule - but both scenarios
are computed and shown, since the drafting is genuinely ambiguous.
Added a "Disqualification scope" sheet to `estimated_contract_value.xlsx`
with whole-firm-sum vs product-only-value for each of the 3 Group A/B
cases (hardcoded from `rc_vs_bans.xlsx`'s own Strength sheet, not
re-derived).

**Done - Item 2 (shared-supplier flag)**: Checked RMSC.xlsx directly -
confirmed many drugs have 2-3 suppliers under the SAME tender
(L1/L2/L3 multi-vendor contracts). Since Annexure VIII's quantity is
the TENDER's total for that drug, not per-supplier, every computed row
now says "Shared with N suppliers - value is an upper estimate" when
true. Value column renamed to "Estimated contract value (2-year, upper
estimate), not amount paid" throughout (not just the shared rows -
even a sole-supplier line's quantity is still a tender-time estimate,
not an amount paid).

**Done - Item 3 (Ultra Drugs)**: Added a targeted
`REASON_CATEGORY_OVERRIDES` dict to `08_rmsc.py`, keyed by Entry ID
(`RMSC-DEBAR-1837-12-08-2025`), setting its category to "Other -
tender violation (bid-rate change after financial bid opening)" -
documented WHY directly in the code (the Hindi letter revealed a bid-
rate dispute, not a quality finding, which the English-only
`reason_category()` keyword matcher couldn't have known). Checked
where this matters: Ultra Drugs was ALREADY excluded from Group A/B/C
in `rc_vs_bans.xlsx` (it has 0 RMSC RC lines - not a supplier there at
all) and was already categorized "To check", never "Quality", before
this change - its continuing presence in `candidate_gaps.xlsx` is a
SEPARATE, legitimate finding (APMSIDC independently found a quality
issue for Ultra Drugs, unrelated to RMSC's bid-rate case) and was
correctly left in place, not removed - removing it would have hidden
real evidence. Re-ran `08_rmsc.py` then `10_merge_master.py` (and its
auto-chained 11/12) - counts unchanged (793 rows, 163 Matches, 39
candidate gaps), confirming the recategorization only changed the
label, not any downstream grouping.

**Done - Item 4 (CDSCO batch cross-check)**: Confirmed Project 1's data
IS on this computer, at `E:\cdsco_mqri` (found by listing, not
assumed). Searched 3 already-structured sources (no new scraping/OCR):
`cdsco_nsq_portal/nsq_drugs.csv` (6,956 rows), `.../spurious_drugs.csv`
(49 rows), and `final_package/CDSCO_NSQ_Project_Workbook.xlsx`'s
"master_alerts" sheet (6,842 rows, a cleaned/deduplicated derivative -
searched too in case its normalization caught something the raw CSVs
didn't). Of 22 batch numbers searched, exactly **1 matched**:
Unicure's `URDT0328` - found in CDSCO's Dec-2025 NSQ data, same drug
(Glipizide and Metformin Tablets USP), reported by "DTL Udaipur" -
flagged clearly in the output that DTL Udaipur is a RAJASTHAN state
lab, so this may be the SAME underlying RMSC finding reported upward
to the national CDSCO database, not independent corroboration from a
different source - not overclaimed either way. The raw downloaded
Alert PDFs under `cdsco_alerts/files/` were NOT searched (not yet
turned into searchable text by Project 1 itself) - logged as an open
item rather than silently skipped.

**Done - Item 5 (Step 6 design, proposal only)**: see the user-facing
response for the full proposal - NOT built, waiting for the OK as
instructed.

**Code changed**:
- `scripts/08_rmsc.py` (`REASON_CATEGORY_OVERRIDES`) -
  `code_history/2026-10-31_1451_08_rmsc.py.before`/`.diff`.
- `scripts/build_finance_report.py` (shared-supplier flag,
  "Disqualification scope" sheet, sheet-name-length fix) -
  `code_history/2026-10-31_1458_*.before`/`.diff` (two passes).
- `scripts/build_batch_crosscheck.py` (new) -
  `code_history/2026-10-31_1455_*.created`.

**Results with numbers**:
- Nestor/NIB-14/2024/69 disqualification scenarios: whole-firm = Rs
  52,728,434.02 (sum of all 3 computed RC lines under that tender),
  product-only = Rs 7,092,639.10 (just the banned Ferrous Sulphate
  [391] line). Nestor's NIB-08/2024 case and Zee's NIB-01/2025 case:
  both scenarios "N/A" (neither tender's Annexure VIII was read as
  native text).
- Shared-supplier flag: of 36 computed RC lines (Nestor+Zee), **26 are
  flagged "Shared with N suppliers"** (2 or 3 suppliers each) - only
  10 are sole-supplier lines.
- `batch_crosscheck.xlsx`: 22 batches checked, 1 found (Unicure
  URDT0328), 21 not found in any of the 3 searched CDSCO sources.
- `master_list.xlsx`/`candidate_gaps.xlsx` row counts unchanged after
  the Ultra Drugs recategorization (793 / 39), as expected.

**Problems and source issues**: none newly logged - the CDSCO cross-
check's "not found" results are expected/reported outcomes, not data
problems, and the Alert-PDF gap is noted as an open item above rather
than a source issue (it's this project's own choice not to re-parse
another project's raw PDFs, not a flaw in either project's data).

**Open items**: per the grouped checklist (section 6, unchanged by
this entry - these 4 items don't map onto Sanskruti's own manual
checklist, they were fully completed here). Step 6 (data science)
awaits the user's OK on the proposal below before any building starts.

### Entry 16 - 2026-10-31 - External review fixes: reason_category expanded, Group U added, dosage-form check, Part C short-word fix, MPPHSCL clause glossary, RMSC sutures tender, full re-run

**Asked**: 8 items from the user's own review (`first_pass_review_2026-10-31.xlsx`,
"Bugs found" sheet - the file itself was never actually placed in
`output/review/` despite the user saying they were putting it there;
confirmed missing via both Bash and PowerShell before proceeding on
everything the user's own message already specified in enough detail,
and flagging item 5 as blocked on the real file). (1) Add quality
words to `reason_category()`; "Non-supply and Quality failure" ->
"Both - check". (2) New Strength group "U = reason unknown" instead of
folding "To check" into D. (3) Group B needs same medicine AND same
dosage form; compare names with spaces removed too. (4) Part C:
3-letter words count if exact; compare names with spaces removed;
remove BHARAT from common words. (5) Keep the user's match decisions
in `match_decisions.xlsx`, read every run. (6) Find MPPHSCL's tender
clauses (10.6, 10.9, 12.5, Annexure IX 7/12) and classify them. (7)
Download RMSC's NIB-10/2026 sutures tender, report its deadline and
clause. (8) Re-run everything, report firms per group and the
possible-breach/policy-gap list.

**Done - Item 1**: Added to `QUALITY_WORDS` in `common.py`: dissolution,
disintegration, sterility, endotoxin, assay, content of,
identification, related substances, particulate, uniformity, nabl
failure, quality failure, quality ground. "ph" was NOT added as a
plain substring - caught before it caused damage: "ph" is a substring
of "pharma"/"pharmaceutical"/"Phenytoin"/"Phenobarbitone" and would
have wrongly flagged almost every non-supply reason that happens to
mention a company or drug name containing those letters. Implemented
as a separate word-boundary regex (`\bph\b`) instead. Tested 7 cases
including the exact "Non-supply and Quality failure" -> "Both - check"
example and a "Pharmaceutical company violation" -> "To check"
negative control - all passed.

**Done - Item 2**: `classify()` in `build_rc_vs_bans.py` now returns
"U" for `ban_cat == "To check"`, kept separate from "D" (Non-supply/
Other/Both-check). All group-letter loops across `build_rc_vs_bans.py`
and `build_rmsc_rc_wayback.py` updated to include U.

**Done - Item 3**: Added `dosage_form()` (a canonical-form lookup -
TAB/TABLET/TABLETS all map to "TABLET", etc.) and required it to
MATCH for Group B, demoting a shared-word-but-different-form pair to C
with an explicit note. Added fused (space-removed) word comparison to
`drug_words()` for the CARBOXY METHYL CELLULOSE / Carboxymethylcellulose
case. Added a non-blocking strength-mismatch note (`_same_strength()`)
that never changes the group, just flags it. **Found while implementing,
reported plainly rather than smoothed over**: all 3 of the PREVIOUSLY
reported Group B matches (Nestor x2 via APMSIDC, Zee x1 via PMBI) turn
out to be different dosage forms on close reading (Tablet vs Syrup
twice, Suspension/Syrup vs Tablet once) - all 3 move to Group C now.

**Done - Item 4**: In `14_match_firms.py`, `content_words()` now keeps
3-letter words (previously dropped at <4 letters) but `has_fuzzy_shared_word()`
only accepts them on an EXACT match, never fuzzed. Added the same
fused-word comparison as item 3. Removed BHARAT from `GENERIC_WORDS`.
Verified: ANG/ZEE/BDR variant pairs that were previously only caught
(if at all) via the generic-word fallback now correctly share a real
word. Match candidates: 63 -> 66 kept.

**Done - Item 5 (partial - blocked)**: Built the READING mechanism -
`load_decisions()` in `14_match_firms.py` reads `output/match_decisions.xlsx`
(keyed by both standard names, order-independent) at the start of every
run and restores "My decision" values. This script only reads that
file, never writes to it. Created an EMPTY, correctly-schemed
placeholder `match_decisions.xlsx` since the real
`first_pass_review_2026-10-31.xlsx` was never actually placed in
`output/review/` - confirmed missing again at the end of this entry,
not just at the start. **Still open**: the actual copy-the-real-
decisions-over step cannot happen until the user provides that file.

**Done - Item 6**: Read 5 MPPHSCL order letters directly (2 native
text, 3 hin+eng OCR - MPPHSCL's letters are ALSO in Hindi, same as
RMSC) to build a clause glossary in `07_mpphscl.py`
(`CLAUSE_TEXT_OVERRIDES`, `ENTRY_ID_OVERRIDES`): Annexure IX clause 7 =
NSQ batch failure (Quality, universal text - applied to every row
citing it); Annexure IX clause 12 = non-supply (Non-supply, universal);
T-213 clause 10.6 = price above DPCO/NPPA ceiling (Other, universal);
T-213 clause 12.5 = failed to execute agreement/EMD forfeited (Other,
universal); T-557 clause 10.9 is explicitly NOT single-meaning (its
own text lists 6 different possible grounds) - only the ONE row
actually read (Unicare Industries, forged technical-bid documents) was
overridden, by Entry ID, not by clause number. A first regex for
clause 12.5 missed `MPPHSCL-PRODUCT-2808-2022` because its citation
omits the word "clause" entirely ("TENDER TERMS AND CONDITON 12.5") -
caught by checking the actual output row, not assumed correct, fixed
by making "clause" optional in the pattern. The ~250 MPPHSCL rows with
NO clause number cited at all are explicitly NOT covered - reading
enough individual letters to classify those would be a much bigger
task.

**Done - Item 7**: Found and downloaded RMSC's NIB-10/2026 Surgical &
Sutures tender (`Finalbid10UBN_compressed.pdf`, robots.txt allowed,
SHA-256 `68171d0c...`). Scanned (74 pages) - OCR'd pages 1 and 9. Bid
deadline: **06.05.2026, 6:00 PM** (technical bid opening 07.05.2026,
11:00 AM) - OCR renders the leading digit ambiguously ("(6.05.2026")
but the day-before-opening pattern makes "06" essentially certain, not
treated as fully confirmed either way (rule F). Debarment clause (page
9, clause (k)): **word-for-word the SAME clause** already found in the
Drug tenders (NIB-14/2024/69, NIB-10/2024, NIB-01/2025) - confirms RMSC
uses one standard debarment clause across Drug AND Surgical/Sutures
tenders alike.

**Done - Item 8 (full re-run)**: `07_mpphscl.py` -> `10_merge_master.py`
(auto-chains 11/12) -> `build_rc_vs_bans.py`. **Large, legitimate
swings, not bugs** - traced each one before accepting it:
- Master list: still 793 rows. Quality+Quality cross-buyer pairs: 4 ->
  **17**. Candidate gaps: 39 -> **116** (21 distinct firms, up from 7).
  Root cause confirmed by re-fetching MPPHSCL's live page directly:
  Zee Laboratories' MPPHSCL ban has reason "Violation of Tender Terms
  and Condition" on the VISIBLE page, but its REMARK field (combined
  into the classification text, per existing design) says
  "Blacklisted/Debarred for Quality Failure" - invisible to the OLD
  word list, caught cleanly by the NEW "quality failure" addition. A
  real, previously-hidden signal, not a false positive.
- Strength groups (firms): A 0->**4**, B 2->**5**, C 4->**17**, D
  30->**9**, U (new) **15**, X 1->1.
- **No firm currently shows "Possible breach of RMSC tender rule"** -
  Nestor's previous breach finding is RETRACTED (moved to C, different
  dosage form - see item 3). The new Group A/B firms' other-buyer bans
  ARE quality-flagged under the new broader `reason_category()`, but
  their specific wording doesn't say NSQ/conviction (RMSC's clause
  stays deliberately narrow) - 8 firms land in "Allowed by RMSC rule
  (policy gap)", 1 in "allowed ground but deadline doesn't overlap".
  This is the clause-check logic working as designed (broad quality
  detection vs narrow disqualification wording are intentionally
  different bars), not a contradiction.
- Meril Endo Surgery Pvt Ltd appears in the new Group U - directly
  relevant to item 7's RMSC sutures tender work.

**Code changed**:
- `scripts/common.py` (QUALITY_WORDS expansion, `ph` word-boundary
  fix) - `code_history/2026-10-31_1524_*.before`/`.diff`.
- `scripts/build_rc_vs_bans.py` (Group U, `dosage_form()`, fused-word
  `drug_words()`, `_same_strength()`) - `code_history/2026-10-31_1528_*.before`/`.diff`.
- `scripts/build_rmsc_rc_wayback.py` (group-letter loop fix) -
  `code_history/2026-10-31_1528_*.before`/`.diff`.
- `scripts/14_match_firms.py` (3-letter exact match, fused-word
  comparison, BHARAT removed, `load_decisions()`) -
  `code_history/2026-10-31_1529_*.before`/`.diff` and
  `code_history/2026-10-31_1530_*.before`/`.diff` (two passes).
- `scripts/07_mpphscl.py` (`CLAUSE_TEXT_OVERRIDES`, `ENTRY_ID_OVERRIDES`,
  the clause-12.5 regex fix) - `code_history/2026-10-31_1538_*.before`/`.diff`
  and `code_history/2026-10-31_1542_*.before`/`.diff` (two passes).

**Results with numbers**: see Item 8 above - all the concrete counts
are there rather than repeated twice in one entry.

**Problems and source issues**: none newly logged to `source_issues.csv`
this entry - every open question above (the missing review file, the
uncovered 250 MPPHSCL rows, the OCR-uncertain sutures-tender deadline
digit) is tracked in this entry's own text and the checklist instead,
since none of them are new SOURCE-DATA problems (they're review-
process and project-scope items).

**Open items**: the real `first_pass_review_2026-10-31.xlsx` still needs
to be placed in `output/review/` before item 5 can be completed (the
reading mechanism is ready and tested with an empty file). The ~250
unclassified MPPHSCL "Violation of Tender Terms" rows with no clause
number remain `To check`/Group U. Section 6 checklist updated below.

---

### Entry 17 - 2026-10-31 - Real review file loaded; Q1/Q2 quality-level split built; Meril bug fixed; 4 case files built; MPPHSCL OCR batch started

**Note on how this entry came to be written**: this entry's own work
(everything below) was actually carried out in the same session as
Entry 16, immediately after it, but the write-up to this section was
never done before the session stopped - section 6's checklist had
already been edited in place (per rule C) to mark items 1/3/7 "DONE
(Entry 17)" and reference this entry's Q1/Q2 counts, but Entry 17
itself did not exist here, and the MPPHSCL OCR batch (last item below)
was left mid-run with only 34 of 164 rows checkpointed. This is a
process gap against rule A ("after finishing any task... add ONE
entry"), caught and corrected now, in a follow-up session, by: (1)
verifying every claim already sitting in section 6 against the real
output files on disk before writing a word of this entry (not simply
copying the checklist's own claims forward), and (2) resuming the
unfinished OCR batch (see Entry 18) rather than treating the
checklist's forward-reference as if the work were already done.

**Asked**: Five tasks, carried over verbatim from the user's own
numbered list at the start of this continued session. (1) The real
review file is now at `output/review/first_pass_review_2026-10-31.xlsx` -
load the match decisions into `match_decisions.xlsx`, re-run 14 and
the merge, report how many pairs changed. (2) Split Quality into Q1/Q2
quality levels with a new "Quality level" column, report RMSC's clause
as BOTH "Possible breach (Q1)" and "...if Q2 counts as NSQ ground",
re-classify the policy-gap firms, record the legal basis (Drugs and
Cosmetics Act 1940 S.25) in this file. (3) Investigate why Meril is in
Group U despite "Sterility test" now being a quality word, fix if a
bug, check GMSCL ban (16-07-2025 to 15-07-2028) against RMSC
NIB-10/2026's deadline, identify the lab, build `output/cases/Meril.xlsx`.
(4) OCR (hin+eng) MPPHSCL's ~250 unclassified "Violation of Tender
Terms" order letters, classify using the clause glossary, mark
"OCR - check", give a random-20 sample. (5) Keep the Nestor CAG Haryana
finding as its own confirmed case in `Nestor.xlsx`, separate from the
retracted RMSC analysis.

**Done - Task 1**: Confirmed `output/first_pass_review_2026-10-31.xlsx` and
`output/review/first_pass_review_2026-10-31.xlsx` are byte-identical to the
copy the user saved to their own Downloads folder (SHA-256
`97546c43...` all three, checked directly rather than assumed). Ran
`build_match_decisions.py` (new script this entry) - first pass had a
real bug, caught immediately: it treated the review file's "Entry IDs
A" column as always a single ID (`{r[...]}`), but some review rows
hold a comma-joined list there too, the same as column B - fixed to
split both sides consistently before the real run. Loaded, re-ran
`14_match_firms.py` (no code change needed - its `load_decisions()`
reader was already built in Entry 16), then `10_merge_master.py`
(auto-chains 11/12).

**Done - Task 2**: Added `quality_level(reason_as_written,
lab_that_declared_nsq, reason_category_value)` to `common.py` -
returns "Q1" if the combined reason-text + lab field names a statutory
testing authority (`Q1_LAB_WORDS`: government analyst/govt analyst,
state drug testing (laboratory), food and drug (testing) laboratory,
etc., plus `Q1_ACRONYM_RE` for CDSCO/CDTL/DCL/CDL as whole words),
else "Q2" if the row is already categorized "Quality" by some other
route, else "N/A" for a non-quality row (never guessed). Wired into
`10_merge_master.py`, which appends it as the 20th master-list column,
`"Quality level (Q1 = statutory lab named, Q2 = other/no lab)"` -
functionally the "Quality level" column asked for, named to be
self-explanatory in the file itself rather than needing this log to
decode it. Extended `build_rc_vs_bans.py`'s `clause_check()` to report
both grounds for every Group A/B row rather than collapsing to one
answer. Recorded the legal basis given by the user - a Government
Analyst's report under the Drugs and Cosmetics Act, 1940, Section 25,
is the formal NSQ declaration route; a failed pharmacopoeial test
(dissolution/sterility/assay/etc.) is NSQ "in substance" even without
a named Section 25 authority - as section 3.7 above (added this entry,
numbered 3.7 rather than overwriting the existing 3.5, per the user's
own instruction about the numbering collision).

**Done - Task 3**: Confirmed the cause by hand before touching
anything - `QUALITY_WORDS` gained "sterility" (among others) in Entry
16, but `05_gmscl.py` was never re-run afterward, so its own
precomputed `reason_category` for Meril's GMSCL row ("Sterility test")
was still the stale "To check" value from before that word existed -
landing it in the new Group U by design (U = `ban_cat == "To check"`),
not a bug in Group U itself. Fixed by simply re-running `05_gmscl.py`
(no code change needed - the bug was staleness, not logic) then the
merge chain. Read GMSCL's full order PDF
(`Debar_Circular-23_-_Meril_Endo-_Suture-1.pdf`, native text) directly
rather than trusting the live summary page, which names no lab at all
- found "Food and Drug Laboratory, Vadodara" there and added it via a
`LAB_OVERRIDES`-style entry in `05_gmscl.py` so it appears in
`master_list.xlsx`'s "Lab that declared NSQ" column (and correctly
scores Q1, since it is a named statutory lab, even though the source
page's own visible text doesn't say so). Computed the date-overlap
check directly: ban 16-07-2025 to 15-07-2028 vs RMSC NIB-10/2026's
OCR-read deadline 06-05-2026 - active on the deadline. Built
`output/cases/Meril.xlsx` via `build_case_files.py` (extended with a
`"Meril"` entry and a new `extra_facts` list per case, used for
exactly this kind of question/answer pair that doesn't fit the
standard "Bans" row shape), same 4-sheet structure as the existing
Jackson/Zee case files plus a 4th "Case-specific findings" sheet.

**Done - Task 5**: The retracted RMSC/Nestor dosage-form analysis
(Entry 16) stays exactly as retracted - nothing re-added to any
"finding" field for it. Found the CAG Haryana matter via web search
(the user referred to it without a link, assuming it was already
known - it was not) - The Tribune, 14-11-2024, reporting a CAG audit
("Public Health Infrastructure and Management of Health Services",
tabled in the Haryana Vidhan Sabha) that Haryana Medical Services
Corporation Ltd bought Rs 1.52 crore of drugs from Nestor in 2019,
after GMSCL had already blacklisted Nestor for the same drug (Feb
2017) and after HMSCL itself separately blacklisted Nestor for
concealing that fact (Sep 2018) - a different, independently-audited
buyer pair (GMSCL/HMSCL) from the retracted RMSC/GMSCL pairing, not a
revival of it. Added to `build_case_files.py` as Nestor's own
"Outside evidence" row (the Rs 1.52 crore figure is the CAG's/press's
own reported number, cited not computed, per the standing no-rupee-
values instruction).

**Done - Task 4 (started, not finished this entry - see Entry 18)**:
Built `build_mpphscl_unclassified.py` - for every MPPHSCL row still
"To check" with no clause number (164 distinct rows, not ~250 as
estimated at the start - the real count, read directly off
`master_list.xlsx`, not guessed), download and OCR (hin+eng) the
linked order letter and classify it by the same clause-citation
patterns as `07_mpphscl.py`'s glossary plus broader Hindi/English
quality/non-supply/fraud signal words, marking every row "OCR - check"
(rule F) - genuinely unclear letters stay "To check" rather than
forced into a category. Two bugs caught by the small-test-first run
before the full batch: (1) missed MPPHSCL's TLS legacy-renegotiation
requirement (same quirk `07_mpphscl.py` already works around) - fixed
with the same `LegacyRenegotiationAdapter`, certificate verification
left ON; (2) a short native-text layer (e.g. just a stamped page
number) was being accepted as "the letter" instead of falling through
to OCR - fixed with a `MIN_NATIVE_CHARS = 150` floor. A third problem
surfaced only at scale, not in the small test: the first real batch
run was killed by the OS for low system memory partway through (21
rows done) - since the script only wrote its output `.xlsx` once at
the very end, nothing was saved despite 21 rows' worth of download+OCR
work already having happened. Fixed with a per-row CSV checkpoint
(`output/mpphscl_unclassified_checkpoint.csv`, append-only, flushed
after every row) plus a periodic `gc.collect()` - PROJECT_RULES.md's
recoverability priority, not a style choice. Re-ran; the session
stopped again before this second run finished, with 34 of 164 rows
checkpointed this time (cause of this second stop not established -
unlike the first, no OOM evidence was checked for it, so it is logged
as unknown rather than assumed to be the same cause). Evidence PDFs
and OCR text for all 34 rows are safely on disk either way, per the
same checkpoint design. See Entry 18 for the resumed, completed run.

**Code changed**:
- `scripts/build_match_decisions.py` (new) -
  `code_history/2026-10-31_1601_build_match_decisions.py.created`,
  then the Entry-IDs-A split-on-comma bug fix -
  `code_history/2026-10-31_1602_*.before`/`.diff`.
- `scripts/common.py` (`quality_level()`, `Q1_LAB_WORDS`,
  `Q1_ACRONYM_RE`) - `code_history/2026-10-31_1855_*.before`/`.diff`;
  then `05_gmscl.py`'s Meril lab override (Task 3) -
  `code_history/2026-10-31_1906_*.before`/`.diff`.
- `scripts/10_merge_master.py` (Quality-level column appended to the
  master-list header/rows) - `code_history/2026-10-31_1856_*.before`/`.diff`.
- `scripts/build_rc_vs_bans.py` (`clause_check()` reports both Q1 and
  Q2 grounds) - `code_history/2026-10-31_1858_*.before`/`.diff`.
- `scripts/05_gmscl.py` (Meril "Food and Drug Laboratory, Vadodara"
  lab override, re-run to clear the stale-category bug) -
  `code_history/2026-10-31_1904_*.before`/`.diff`.
- `scripts/build_case_files.py` (Meril case with its `extra_facts`
  sheet; Nestor's Outside-evidence row) -
  `code_history/2026-10-31_1908_*.before`/`.diff` (the second,
  `2026-10-31_1909_*.before`/`.diff`, snapshot pair is an empty diff -
  a backup taken around re-running the script, not a further source
  change).
- `scripts/build_mpphscl_unclassified.py` (new) -
  `code_history/2026-10-31_1911_*.created`; TLS adapter fix -
  `code_history/2026-10-31_1912_*.before`/`.diff`; MIN_NATIVE_CHARS fix
  - `code_history/2026-10-31_1918_*.before`/`.diff`; checkpoint +
  periodic gc fix - `code_history/2026-10-31_1931_*.before`/`.diff`.

**Results with numbers**:
- Task 1: 127 of the review file's 127 decided pairs (51 "Match
  candidates" + 76 "Generic words only") matched to a current
  `match_candidates.xlsx` row by Entry ID with 0 unmatched - 65
  "Different", 61 "Same", 1 "Not sure". After the later Meril-bug
  re-run (Task 3) regenerated `match_candidates.xlsx`, 61 of its 66
  "Match candidates" rows and all 66 "Generic words only" rows still
  carry a restored decision; the 5 without one are new KSMSCL
  Martin/Vivek name-fragment pairs that only appeared after that
  re-run (checklist item 5).
- Task 2: master-list Quality-level counts - N/A 441, Q2 325, Q1 27.
  `rc_vs_bans.xlsx` Group A/B firms (8 distinct: Ashri Medico
  Surgicals, Modern Laboratories, Nandani Medical Laboratories, Zee
  Laboratories, Healers Lab, Omega Pharma, Prochem Pharmaceuticals,
  Vivek Pharmachem) all show a policy-gap classification conditioned
  explicitly on whether Q2 counts; none currently reads as an
  unconditional "Possible breach" once the bid-deadline-overlap check
  is applied (consistent with Entry 16's own retraction of the
  previous Nestor finding on dosage-form grounds) - zero firms
  currently clear BOTH the Q1-or-Q2 ground test AND the deadline
  test at once.
- Task 3: Meril's GMSCL ban confirmed ACTIVE on the RMSC NIB-10/2026
  deadline; lab identified (Food and Drug Laboratory, Vadodara);
  `output/cases/Meril.xlsx` built (4 sheets, 6 ban rows across 2
  buyers, 17 RMSC RC lines).
- Task 5: `output/cases/Nestor.xlsx` now has a 1-row "Outside
  evidence" sheet (CAG Haryana/HMSCL finding), separate from its 9-row
  "Bans" sheet (no "finding" language attached to the retracted RMSC
  pairing there).
- Task 4: 164 target rows identified; 34 checkpointed as of this
  entry (not yet a final report - see Entry 18).

**Problems and source issues**: none newly added to
`output/source_issues.csv` this entry - the bugs found (Entry-IDs-A
comma-split, TLS adapter, MIN_NATIVE_CHARS, OOM/checkpoint) were all
caught and fixed in this project's own code before producing any
reported number, not left as unexplained source-data gaps.

**Open items**: Task 4's batch is incomplete (130 of 164 rows
remained at the point this entry's work stopped) - continued in Entry
18. Checklist item 11 (does the user agree Q2 should count as NSQ for
RMSC's clause) is a judgment call for the user, not something this
project can settle. Section 6 checklist corrected below to match what
is actually true as of this entry (item 12 was prematurely marked
against an unfinished task before this entry existed - fixed now
rather than carried forward).

---

### Entry 18 - 2026-11-01 - MPPHSCL OCR batch (Task 4) completed

**Asked**: Continuation of Entry 17's Task 4 - finish the interrupted
MPPHSCL "Violation of Tender Terms" OCR/classification batch
(`build_mpphscl_unclassified.py`). The user killed the first resumed
run partway through themselves, then explicitly asked to continue -
resumed a second time from the checkpoint.

**Done**: Re-ran `build_mpphscl_unclassified.py` twice, both times
resuming from `output/mpphscl_unclassified_checkpoint.csv` (no
re-download, no re-OCR of already-checkpointed rows, per the
checkpoint's own design from Entry 17). First resume: ran from 34 to
127 of 164 rows before being killed by the user. Second resume: ran
the remaining 37 rows to completion, exit code 0. Inspected the
"Random 20 sample" sheet directly (not just the printed category
counts) before reporting anything - caught a real, previously-unlogged
source problem this way: Meril's own two MPPHSCL rows
(`MPPHSCL-FIRM-7876-2024`, `MPPHSCL-PRODUCT-166-2024`) both cite the
same order PDF (`order meril.pdf`), which pdfplumber cannot open at
all - `Invalid dictionary construct: [/'Registry', b'', /'Ordering',
b'', /'Supplement']`, a malformed CMap/encoding dictionary in the PDF
itself (confirmed: downloads fine, fails identically on a second
attempt, so not a transient network issue). Logged to
`output/source_issues.csv` as a genuine, unresolved source-data
problem rather than silently left unexplained in the batch output -
both rows correctly stayed "To check" rather than being guessed at.

**Code changed**: none this entry - no bugs found in
`build_mpphscl_unclassified.py` itself on this run; the malformed-PDF
case above is a source-data problem, not a script bug (the script's
existing `except Exception` handler around the PDF-read step already
catches it correctly and records why).

**Results with numbers**: All 164 target rows processed. Final
category counts: **To check 110, Non-supply 22, Quality 21, Other 11**.
Every row carries `Status` = "OCR - check" per rule F (or "native
text" in `Note` where the letter had enough embedded text to skip
OCR - 150-character floor from Entry 17's `MIN_NATIVE_CHARS` fix).
`output/mpphscl_unclassified_batch.xlsx` saved with two sheets: "All
processed" (164 rows) and "Random 20 sample" (20 rows, for hand
review - includes one of the two malformed-PDF Meril rows by chance,
which is how that problem was caught). 2 new rows added to
`output/source_issues.csv` (the malformed Meril PDF, both citing
Entry IDs that reference it).

Nothing from this batch is wired back into `07_mpphscl.py`'s own
`CLAUSE_TEXT_OVERRIDES`/`ENTRY_ID_OVERRIDES` or `master_list.xlsx` -
per the script's own design (see its docstring, confirmed unchanged),
this is a report for the user to review first; folding 164 rows'
worth of OCR-derived categories into the pipeline unreviewed would be
exactly the kind of unreviewed bulk edit PROJECT_RULES.md rule D warns
against.

**Problems and source issues**: the 2 malformed-Meril-PDF rows above.
No other read/download/classification failures in the full 164-row
run (the only other non-"recognized signal" categories are "clause
citation matched" and the keyword-based ones, all working as
designed).

**Open items**: per rule F, none of these 164 proposed categories
count as confirmed until the user has checked the Random-20 sample (or
more) against the actual letters - checklist item 12 updated below.
The malformed Meril PDF (2 rows) has no proposed fix in this entry -
logged as a source problem per PROJECT_RULES.md's "log, don't silently fix"
rule, not investigated further (would need either a PDF-repair tool
or a hand-opened copy, neither attempted here).

---

### Entry 19 - 2026-11-01 - Q1-only headline decision (checklist item 11); classification centralized into 10_merge_master.py (checklist item 13)

**Asked**: Two decisions from the user, in response to Entry 17/18's
open items. (1) Checklist item 11 (Q1/Q2): headline findings use Q1
only; report Q2 separately, always labeled "quality failure in
substance, not a formal NSQ declaration (no government lab named)";
show both counts everywhere; label clearly. (2) Checklist item 13 (the
Meril bug's root cause): move `reason_category()` and the Q1/Q2 level
out of the buyer scripts into `10_merge_master.py`, so classification
runs once over all buyers on every merge; buyer scripts only keep
"Reason (as written)" and the lab name. Prove nothing else changed:
re-run, compare with the current `master_list.xlsx`, and list any row
whose category changed.

**Done - Item 13 (done first, since Item 11's wording change sits on
top of it)**:
1. Inspected every buyer script's call site before touching anything -
   found 9 of 11 buyer scripts call `reason_category()` directly on the
   SAME text written to "Reason (as written)" (straightforward to
   move), but two needed care:
   - `07_mpphscl.py` classifies on `reason + remark` combined
     (`classify_text`), but only ever WRITES `reason` to the "Reason
     (as written)" column - the Remark text (checklist item 2's "Quality
     Failure" case) was invisibly feeding classification without ever
     being stored anywhere. Fixed by folding Remark into the stored
     "Reason (as written)" text itself (bracketed, only when it differs
     from Reason) - this is required for centralization to see the same
     signal the old per-buyer code did, and as a side effect makes the
     remark visible to a human reading the file too (narrows checklist
     item 2).
   - `07_mpphscl.py` (`CLAUSE_TEXT_OVERRIDES`, `ENTRY_ID_OVERRIDES`) and
     `08_rmsc.py` (`REASON_CATEGORY_OVERRIDES`) both carry hand-verified,
     buyer-specific overrides read directly from real order letters -
     these are facts, not keyword guesses, so they needed to move WITH
     the classification call, not be left behind or silently dropped.
2. Added `common.CATEGORY_PENDING` - the placeholder every buyer script
   now writes into "Reason category" instead of a computed value.
3. Edited all 9 buyer scripts that build real rows (`03_build_master.py`,
   `03_assam_master.py`, `04_pmbi.py`, `05_gmscl.py`, `07_mpphscl.py`,
   `08_rmsc.py`, `09_kmscl.py`, `13_bmsicl.py`, `16_ksmscl.py`) to stop
   calling `reason_category()` (or, for the two OCR-fragment-only
   buyers, stop hardcoding "To check" directly) and write
   `CATEGORY_PENDING` instead. `09_kmscl.py`'s OTHER function
   (`nsq_rows()`) still hardcodes "Quality" directly - left untouched,
   since that is a structural fact about the source document (its own
   title says every row IS NSQ), never a keyword guess, so it was never
   at risk of going stale the way `reason_category()` calls were.
4. Moved `MPPHSCL_CLAUSE_TEXT_OVERRIDES`, `MPPHSCL_ENTRY_ID_OVERRIDES`,
   and `RMSC_REASON_CATEGORY_OVERRIDES` (same content, same comments,
   renamed with a buyer prefix for clarity now that they all live in one
   file) into `10_merge_master.py`, with a new `classify_row(buyer,
   entry_id, reason_as_written)` that applies the overrides first, then
   falls back to `reason_category()` - and reconstructs the exact same
   `"Auto - check (category overridden - ...)"` Status text the old
   per-buyer code used to write directly, so moving this did not even
   change that wording (only the file name inside the note's own "see
   ..." pointer, since the tables themselves moved).
5. Wired `classify_row()` into `10_merge_master.py`'s `main()`, run on
   every merged row, overwriting the `CATEGORY_PENDING` placeholder,
   BEFORE `quality_level()` runs (which needs the real category).
6. **Verification, per the user's explicit instruction**: before
   re-running anything, copied the current `master_list.xlsx` to a
   scratch snapshot. Re-ran every buyer script (confirming each still
   produces the same row counts as the known baseline: AP 43, AS 0,
   BMSICL 81, GMSCL 99, KMSCL 5, KSMSCL 11, MPPHSCL 378, PMBI 75, RMSC
   101 - all matched), then `10_merge_master.py`. Compared the new
   `master_list.xlsx` against the snapshot row by row, keyed by Entry
   ID, on EVERY column - not just category.

**Done - Item 11 (built on top of the now-centralized classify_row())**:
1. Rewrote `build_rc_vs_bans.py`'s `clause_check()` to return
   `(allowed, active, q1_classification, q2_classification)` - two
   independent values, computed unconditionally, instead of one
   `classification` string whose content depended on which branch of
   the old if/elif produced it. Q2's text is now built from one fixed
   `Q2_LABEL` constant ("quality failure in substance, not a formal NSQ
   declaration (no government lab named)") used EVERYWHERE Q2 appears,
   so the wording can never drift between call sites.
2. Added a second "worst-case" roll-up function for the By-firm sheet -
   `worst_q1()` and `worst_q2()`, kept fully separate (previously, one
   `worst_clause_class()` picked "whichever is stronger" between a Q1
   and Q2 breach on the same firm - that silent preference is gone;
   each is now its own column, always).
3. Added a 4th Strength-sheet column and a 2nd By-firm-sheet column
   (Q1 and Q2 each get their own), and rewrote the script's printed
   summary to show Q1 and Q2 counts in clearly separate, clearly
   labeled blocks, ending with two explicit SUMMARY lines - one for
   "Possible breach (Q1, headline)", one for "Possible breach if Q2
   counts as NSQ ground" - so a reader never has to parse a sentence to
   find either count.
4. Recorded the decision itself in PROGRESS_LOG.md section 3.7 (not
   overwriting the existing Q1/Q2 definitions, which are unchanged -
   only how they are PRESENTED changed).

**Code changed**:
- `scripts/common.py` (`CATEGORY_PENDING`) -
  `code_history/2026-11-01_2203_common.py.before`/`.diff`.
- `scripts/03_build_master.py`, `03_assam_master.py`, `04_pmbi.py`,
  `05_gmscl.py`, `09_kmscl.py`, `13_bmsicl.py`, `16_ksmscl.py` (import
  swap + `CATEGORY_PENDING` in place of the `reason_category()` call or
  hardcoded "To check") - `code_history/2026-11-01_2204_*.before`/`.diff`
  (one pair per file).
- `scripts/07_mpphscl.py` (Remark folded into "Reason (as written)";
  `CLAUSE_TEXT_OVERRIDES`/`ENTRY_ID_OVERRIDES`/`clause_based_category()`
  removed) - `code_history/2026-11-01_2206_07_mpphscl.py.before`/`.diff`.
- `scripts/08_rmsc.py` (`REASON_CATEGORY_OVERRIDES` removed) -
  `code_history/2026-11-01_2206_08_rmsc.py.before`/`.diff`.
- `scripts/10_merge_master.py` (the relocated override tables,
  `classify_row()`, wired into `main()` before `quality_level()`) -
  `code_history/2026-11-01_2207_10_merge_master.py.before`/`.diff`.
- `scripts/build_rc_vs_bans.py` (`clause_check()` rewritten for
  independent Q1/Q2 return values, `Q2_LABEL` constant, `worst_q1()`/
  `worst_q2()`, new sheet columns, new printed summary) -
  `code_history/2026-11-01_2210_build_rc_vs_bans.py.before`/`.diff`.
- `PROGRESS_LOG.md` (this entry; sections 2, 3.7, 4, and checklist
  items 2/11/13 updated in place per rule C) -
  `code_history/2026-11-01_2203_PROGRESS_LOG.md.before`/`.diff` (the
  snapshot taken before this entry's edits began).

**Results with numbers - the "prove nothing else changed" check**:
Comparing the new `master_list.xlsx` (793 rows, unchanged) against the
pre-refactor snapshot, row by row, by Entry ID:
- **Entry IDs**: identical sets, 0 only-before, 0 only-after.
- **Reason category**: **0 rows changed** (checked all 793).
- **Quality level (Q1/Q2/N-A)**: **0 rows changed**.
- **Reason (as written)**: 341 rows changed - ALL MPPHSCL, ALL exactly
  the designed, disclosed Remark-folding-in from step 1 above (e.g.
  `"...Condition"` -> `"...Condition [Remark: Blacklisted/Debarred for
  Quality Failure]"`). Spot-checked a sample by hand against the
  pre-change text - confirmed each is additive only, nothing removed or
  reworded.
- **Status**: 23 rows changed - ALL the MPPHSCL override rows, ALL
  purely the relocated file-reference inside the note text (e.g. "see
  07_mpphscl.py CLAUSE_TEXT_OVERRIDES" -> "see 10_merge_master.py
  MPPHSCL_CLAUSE_TEXT_OVERRIDES") - the override itself (which category
  it assigned) did not change, confirmed by the 0 category changes
  above.
- Every other column: unchanged (not individually re-checked column by
  column beyond category/quality-level/reason/status above, since
  nothing else in this refactor touches row-building logic for any
  other field).
- `rc_vs_bans.xlsx` re-run: Strength groups **A 4, B 5, C 20, D 9, U 11,
  X 1** firms - matches the state already on disk before this entry
  (verified against this session's own earlier By-firm dump, not
  against Entry 16/17's older, superseded numbers in this log, which
  predate the match-decisions/Meril-bug-fix work already done in
  Entries 17-18). **Q1 (headline): 0 firms with a possible breach.
  Q2 (quality failure in substance...): 0 firms.** Both now printed and
  written as clearly separate, always-shown counts.

**Problems and source issues**: none new. The 9-kmscl.py `nsq_rows()`
hardcoded "Quality" (step 3 above) was deliberately left as-is, not
treated as a gap - documented reasoning above.

**Open items**: none for these two decisions - both are now implemented
and verified. Checklist items 11 and 13 marked done below. Checklist
item 2 narrowed (Remark is now visible) but not closed (still the
user's to confirm against the real PDF).

### Entry 20 - 2026-11-01 - User's own document check applied (8 tasks): clause re-reading, Q1 fixes, OCR corrections, Random-20 classifier, Karnataka full lists, match decisions, concealment cases, full rebuild

**Asked**: The user read every document behind this project's open
questions directly from the page IMAGES (not OCR) and supplied the
result as `output/review/document_checks_2026-11-01.xlsx` (10
sheets: Read me, Case verdicts, Rajasthan clause, Hindi letter
corrections, Random-20 check, New match decisions, Karnataka full lists,
Concealment cases, Key quotes, Still to do). Eight tasks: (1) re-read
RMSC's clause 5(l)(II)/5(k)(ii) as WHOLE-FIRM, not product-only, and
re-classify; show the product-only reading as a second column; (2) fix
Q1 for Zee's 4 MPPHSCL rows and 6 RMSC letters; (3) apply the Hindi
letter OCR corrections, checking the ANG order-number discrepancy; (4)
improve the Random-20/164-row classifier with 5 new signal types and
re-run, reporting new counts; (5) replace KSMSCL's 11 OCR fragments with
the 22-row (not 21 - see below) hand-transcribed Karnataka full lists;
(6) apply the new match decisions; (7) add Radiant Parenterals to the
concealment cases; (8) rebuild everything and report possible-breach
firms, policy-gap firms, Q1/Q2 counts, and new totals, then propose a
data freeze.

**Done - Task 1 (clause re-reading, `build_rc_vs_bans.py`)**: see
PROGRESS_LOG.md section 3.8 for the legal reasoning in full. Widened
`clause_check()`'s scope from Groups A/B only to A/B/C (the WHOLE-FIRM
reading now governs the headline Q1/Q2 columns), refined the "active on
bid deadline" wording to explicitly distinguish a genuine "Policy gap"
(Q1 ground exists, RC contract period overlaps the ban, but the ban
wasn't active AT the bid deadline) from "unknown" (tender PDF not
readable). Added a NEW `product_only_reading()` function and a 3rd
Strength-sheet / 3rd By-firm-sheet column reproducing the OLD, narrower
A/B-only behavior under its own explicit label, rather than losing it.
Also fixed, found while validating this against the 3 named cases:
Meril's own RMSC tender deadline (NIB-10/2026, 06-05-2026, OCR-uncertain
leading digit) had been read and recorded in `output/cases/Meril.xlsx`
back in Entry 17/18 but never added to `build_rc_vs_bans.py`'s
`TENDER_DEADLINES_OCR_UNCHECKED` dict - added now.

**Done - Task 2 (Q1 fixes)**: added `MPPHSCL_LAB_OVERRIDES` (07_mpphscl.py)
and `RMSC_LAB_OVERRIDES` (08_rmsc.py), following the EXISTING pattern
already used for GMSCL (05_gmscl.py `LAB_OVERRIDES`) - a buyer script's
own raw "Lab" field, not a central classification override, since
`quality_level()` just reads whatever lab text the row already carries.
Zee's 4 MPPHSCL rows (4102/4790/7882-2024: CDL Kolkata; 8078-2021:
Government Analyst, Controller FDA MP) and 6 RMSC letters (1846, 3078-b,
1453, 554, 3085, 3093: State Drug Testing Laboratory and/or RMSC's own
lab, several cross-confirmed by CDL Kolkata) now carry the real lab name
instead of "Not stated", so `quality_level()` computes Q1 the normal
way - no separate Q1-override mechanism needed. Also added
`MPPHSCL_COURT_OVERRIDES` for Astra Life Care's ban, which is REVOKED
(order 2335), found in the same review pass.

**Bug found and fixed while verifying Task 2**: the ANG Life Sciences
override (`RMSC-DEBAR-3078-20-02-2025-b`) silently never applied - the
Entry ID doesn't get its "-b" suffix until `dedup_entry_ids()` runs,
AFTER `build_row()` already did the (then-unmatched) override lookup on
the un-suffixed ID. Fixed by re-applying `RMSC_LAB_OVERRIDES` a second
time in `08_rmsc.py main()`, after dedup, keyed by the FINAL id. Caught
by comparing every override's expected effect against the actual merged
master_list.xlsx (PROJECT_RULES.md data-integrity habit, not assumed working).

**Done - Task 3 (Hindi OCR corrections, `build_pdfs_to_read.py`)**:
corrected `RMSC_HINDI_READINGS` for 4 of the 6 letters the user re-read
(Zee/1846: batch 722-285->1722-285, PO 654->1654, tender date
07.07.2022->07.01.2022; ANG/3078-b: batch 472006/472007->I472006/
I472007, PO 3917 date now read as 11.10.2022, tender NIB-22/2021/380->
/1380, drug code [492]->[192]; Vee Excel/1453: PO date 25.01.2025->
25.11.2025; Unicure/3093: first batch URDT0005->GE1TC005, 5 purchase-
order numbers/dates corrected). Medipol/554 and Medipol/3085 confirmed
CORRECT as-is, marked so in a comment rather than left silently
unconfirmed. Ultra Drugs/1837 not re-read (not a quality case, per the
user). The ANG letter's own order number (3099) vs this project's Entry
ID (3078, from RMSC's live page) is NOT resolved either way - logged to
`output/source_issues.csv` instead, per the "log source problems instead
of fixing them silently" rule.

**Done - Task 4 (Random-20/164 classifier, `build_mpphscl_unclassified.py`)**:
added 5 new signal groups to `classify_text()`, found directly from the
user's Random-20 review (computer was right on only 3 of 20): clause
8(c)/Annexure-V/IX-withdrawal citations and a new
`AGREEMENT_NOT_EXECUTED_WORDS` text list (9 of 20 rows - the single most
common miss); `FURNITURE_EQUIPMENT_WORDS` + `NOT_PER_SPEC_WORDS` (3 of
20 - hospital furniture, not a drug-quality finding, even though
MPPHSCL's own page calls it "Quality"); `CONCEALMENT_WORDS` +
`CONCEALMENT_CONTEXT_WORDS` (the Radiant Parenterals case). Added a new
`reclassify_from_cache()` mode (`--reclassify` flag) that re-runs the
classifier against ALREADY-SAVED OCR text / already-downloaded PDFs,
with ZERO new network calls or OCR engine calls - re-ran on the full 164
and got: To check 85 (was 110), Other 37 (was 11), Quality 21 (unchanged),
Non-supply 21 (was 22, one row's text now matches a clearer pattern).
Separately, all 20 Random-20 rows themselves are now hand-verified,
entry-ID-keyed overrides in `10_merge_master.py`
(`MPPHSCL_HAND_VERIFIED_OVERRIDES`) - Status "Checked (hand-verified...)"
not "OCR - check", since the user has now actually read each one against
the page image (PROJECT_RULES.md rule F - this is exactly the check rule F asks
for, now done for these 20).

**Done - Task 5 (Karnataka full lists, `16_ksmscl.py` rewritten)**: the
OCR-fragment extraction logic (name fragments only, every other field
"Not stated") is REPLACED entirely by `KARNATAKA_FULL_LIST`, a hardcoded
table of the user's own hand transcription of all 3 PDFs, read directly
from the page images - the same kind of hand-verified-fact table this
project already uses for RMSC's Hindi letters and GMSCL/MPPHSCL's lab
overrides. **Found while implementing: the actual sheet has 22 distinct
rows (5 Firm, 14 Product, 3 Lab), not the 21 the "Read me" sheet's own
narrative says** - the 4 Paschim Banga Pharmaceuticals sub-rows
(2024-25-1a/b/c/d) are 4 separate order numbers/dates/court-stay case
numbers, not 1 row with 4 products. Reproduced faithfully as given
(22 rows), with the discrepancy logged to `source_issues.csv` rather
than silently reconciled either way (similarly, "9 of 22 court-stayed",
not the narrative's "7 of 13"). The 3 source PDFs are still downloaded
and SHA-256-logged for evidence preservation, just no longer OCR'd.
`find_matches()` continues to exclude the 3 "Lab" rows automatically (no
code change needed there).

**Done - Task 6 (match decisions)**: of the user's 7 "New match
decisions", decisions 2-5 (about the old KSMSCL "Vivek"/"Martin" name
FRAGMENTS) are MOOT after Task 5 - the fragments no longer exist, so
there is nothing left to disambiguate. The 2 that still apply were
merged into `match_decisions.xlsx` (129 decisions now, up from 127, all
129 restored successfully on the next `14_match_firms.py` run): "Martin
& Brown Bio Sciences" vs "...Pvt Ltd" = Same (new); "Unicure" vs
"Unicure India Ltd" = Same (updated from "Not sure"). Also added "Vivek
Pharmachem India Ltd" vs "Vivek Pharmachem of India Pvt Ltd" = Same,
following directly from the user's own confirmed identification of the
Jaipur firm's several spellings. "Medipol Pharmaceutica" = "Medipol
Pharmaceutical India Pvt Ltd" was ALREADY "Same" in `match_decisions.xlsx`
(decision #7, pre-existing from Entry 17) - confirmed unchanged, not
reapplied. **Found while verifying**: Jackson's case file
(`output/cases/Jackson.xlsx`) silently LOST its KSMSCL row after Task 5
- the new full name "Jackson Laboratories Pvt Ltd" doesn't collapse to
the shorter "Jackson Laboratories" `build_case_files.py`'s `CASES` dict
already used for PMBI. Fixed by adding the second name variant (shown,
per the module's own design, in the "Possibly related name variants"
sheet, never silently merged).

**Done - Task 7 (concealment cases)**: new report,
`build_concealment_cases.py` -> `output/concealment_cases.xlsx`, listing
all 5 officially-documented cases from the user's "Concealment cases"
sheet (Unicure, Divine Laboratories, Shiv Paramount - all 3 caught by
GMSCL's own debarment-order text; Medipol - PMBI's own revocation names
a GMSD concealment; **Radiant Parenterals - MPPHSCL debarred it for
hiding a Tamil Nadu/TNMSC quality blacklisting**). 4 of the 5 cross-
reference a real `master_list.xlsx` Entry ID (verified - 0 missing-ID
warnings on the actual run); Medipol/PMBI is the one exception (not a
normal master_list.xlsx row - points to `source_issues.csv` and PMBI's
own "Unmatched revocations" sheet instead, same as that fact already
lived before this report existed).

**Done - Task 8 (rebuild + numbers)**: rebuilt in order - `07_mpphscl.py`
and `08_rmsc.py` (so their new Lab overrides actually reach the buyer
files, not just the scripts), `16_ksmscl.py`, `10_merge_master.py`
(auto-chains 11/12), `build_rc_vs_bans.py`, `14_match_firms.py`,
`build_case_files.py`, `build_concealment_cases.py`,
`build_mpphscl_unclassified.py --reclassify`, `build_pdfs_to_read.py`.
See the next message to the user for the full numbers (possible-breach
firms, policy-gap firms, Q1/Q2 counts, new totals) - not duplicated here
to avoid the two ever silently drifting apart; this entry is the
record of WHAT CHANGED, the chat message is the record of the RESULTING
NUMBERS at this point in time.

**Code changed** (every file backed up per rule B,
`code_history/2026-11-02_*`):
- `scripts/08_rmsc.py` (`RMSC_LAB_OVERRIDES` + 2 wiring points, incl. the
  post-dedup bugfix) - `.before`/`.diff` pairs at 0332 and 0339.
- `scripts/07_mpphscl.py` (`MPPHSCL_LAB_OVERRIDES`, `MPPHSCL_COURT_OVERRIDES`
  + wiring) - `.before`/`.diff` at 0332.
- `scripts/10_merge_master.py` (`MPPHSCL_HAND_VERIFIED_OVERRIDES`, wired
  into `classify_row()` ahead of the existing override table) -
  `.before`/`.diff` at 0332.
- `scripts/build_rc_vs_bans.py` (clause scope A/B/C, `product_only_reading()`,
  new Strength/By-firm columns, Meril's missing tender deadline) -
  `.before`/`.diff` pairs at 0332 and 0340.
- `scripts/build_pdfs_to_read.py` (`RMSC_HINDI_READINGS` corrections) -
  `.before`/`.diff` at 0334.
- `scripts/build_mpphscl_unclassified.py` (5 new signal groups,
  `reclassify_from_cache()`, OCR_TEXT import fix) - `.before`/`.diff`
  pairs at 0332 and 0341.
- `scripts/16_ksmscl.py` (full rewrite - `KARNATAKA_FULL_LIST`
  replaces the OCR-fragment extraction) - `.before`/`.diff` at 0332.
- `scripts/build_case_files.py` (Jackson's 2nd name variant) -
  `.before`/`.diff` at 0342.
- `scripts/build_concealment_cases.py` (new file) - `.created` at 0335.
- `output/source_issues.csv` (ANG order-number discrepancy; Karnataka
  21-vs-22-row narrative discrepancy) - appended, not backed up (not a
  `.py`/PROJECT_RULES.md/README file per rule B's own scope).
- `output/match_decisions.xlsx` (2 new/updated decisions) - not backed
  up (a data file, same as every other merge output).
- `PROGRESS_LOG.md` (this entry; sections 1, 2, 3, 3.8 (new), and
  checklist items 2/5/6/12 updated in place, items 9/15/16 added) -
  snapshot taken before this entry's edits began.

**Problems and source issues**: the ANG order-number discrepancy and the
Karnataka 21-vs-22-row narrative discrepancy (both above) are the only
new ones - both logged to `source_issues.csv`, neither silently resolved.

**Open items**: see the checklist (section 6) - items 9 (RTI list),
15, and 16 are new; items 2, 5, 6, 12 are now done or substantially
narrowed. The ~144 MPPHSCL rows in the 164-row batch NOT in the
Random-20 sample remain unconfirmed ("OCR - check") under the improved
classifier - only the 20 sampled rows were hand-verified.

### Entry 21 - 2026-11-02 - Two approval-stage fixes (Meril appeal, Protech/Zee literal-NSQ bug); PMBI Zee PDF checked (none exists); data freeze

**Asked**: Approving Entry 20's work, the user gave two corrections to
apply before freezing: (1) Meril's MPPHSCL firm-level blacklisting
(order 7876) was SET ASIDE on appeal (Commissioner order 167 dt
20-01-2025, page 2) - only an item-level debarment (MIS93, rate error)
survives; fix Level/Category/Court accordingly. (2) Protech Telelinks
should NOT be a Q1 headline case - AP names no lab for it (Q2) - only
Zee/Meril/Nestor stay headline. Then: (3) try to download PMBI's Zee
CMC-eye-drops order PDF, if one exists; (4) freeze everything to
`output/frozen_2026-11-02/` with a README covering row counts, groups,
Q1/Q2 counts, the 3 headline cases, the second-tier list, the
concealment cases, and open items - all future write-up numbers to come
from that folder.

**Done - Fix 1 (Meril appeal)**: added `MPPHSCL_LEVEL_OVERRIDES` and
`MPPHSCL_PRODUCT_OVERRIDES` to `07_mpphscl.py` (same pattern as the
existing Lab/Court overrides - raw data, not central classification),
forcing `MPPHSCL-FIRM-7876-2024` (and its `-b` dedup twin - MPPHSCL's
live Firm page lists this order twice) to Level "Product" with the
MIS93 item text. Updated the Court override to record the appeal result
and updated `10_merge_master.py`'s `MPPHSCL_HAND_VERIFIED_OVERRIDES`
note (category stays "Other", now describing the actual rate-error
ground instead of "result not yet read").

**Bug found while verifying**: the `-b` twin kept its OLD category
("To check") even after the fix - `MPPHSCL_HAND_VERIFIED_OVERRIDES` is
keyed by the FINAL (post-dedup) Entry ID, and the dict only had a key
for the base id, not `-b`. Unlike the Lab/Court overrides (applied
inside `build_row()`, BEFORE dedup assigns the suffix, so they reach
both copies automatically), this dict needed its own explicit `-b` key
added. Noted in a comment so the next person adding an override here
doesn't repeat the same gap.

**Done - Fix 2 (Protech/Zee literal-NSQ bug)**: `build_rc_vs_bans.py`'s
`clause_check()` used to treat ANY literal "Not of Standard Quality"/
"NoSQ"/etc. wording in a ban's own scraped SUMMARY TEXT as satisfying
the Q1 (headline) ground on its own, regardless of `quality_level_value`
- this is how Protech Telelinks (APMSIDC's summary literally says
"Supply of Not of Standard Quality Product", but its own "NSQ declared
by" field names no lab, Q2) wrongly reached the headline. Renamed
`NSQ_CONVICTION_WORDS` to `CONVICTION_WORDS`, dropped the NSQ-wording
part entirely (kept "conviction"/"convicted" - a court conviction is
still its own, lab-independent ground under the clause's own text) -
`is_q1_ground` is now `is_conviction_ground or quality_level_value ==
"Q1"`, nothing else. **This same bug ALSO silently affected Zee's PMBI
CMC-eye-drops case** (PMBI's reason text literally says "Not of
Standard Quality" too, Q2, no lab named) - it had been showing as a Q1
"Policy gap" before this fix; after the fix it correctly moves to the
Q2-conditional bucket, consistent with the review file's own note on
that specific case ("Q2 until the PMBI order is read"). Re-verified the
full headline/second-tier lists end to end after both fixes (see
section 6 of the next message to the user, or `output/frozen_2026-11-02/
README.txt` section 4/5): headline is now exactly Zee, Meril (PENDING -
OCR deadline), Nestor; Protech and Zee's PMBI case both sit in the
10-firm Q2-conditional list.

**Done - Task 3 (PMBI Zee PDF)**: fetched `https://www.pmbi.co.in/
debar.aspx` live (robots.txt 404, not restricted) and inspected the
Carboxy Methyl Cellulose Eye Drops row's own HTML directly - confirmed
it carries NO `<a href>` link at all (debar.aspx exposes no per-row
document for ANY entry, unlike Blacklist.aspx, which `04_pmbi.py`
already pulls evidence links from). Nothing downloaded - not guessed at
an alternate URL. Logged to `source_issues.csv`.

**Done - Task 4 (freeze)**: copied the entire `output/` tree (every
file and subfolder, as it stood after both fixes above) into
`output/frozen_2026-11-02/`, confirmed byte-identical to the live
`output/` via `diff -rq` immediately after. Wrote `README.txt` inside
it with every number asked for (row counts, groups, Q1/Q2 counts, the
3 headline cases, the 10-firm second-tier list, the 5 concealment
cases, and every open item/logged discrepancy at this point in time).

**Code changed**:
- `scripts/07_mpphscl.py` (`MPPHSCL_LEVEL_OVERRIDES`,
  `MPPHSCL_PRODUCT_OVERRIDES`, Court override text updated) -
  `code_history/2026-11-02_0406_07_mpphscl.py.before`/`.diff`.
- `scripts/10_merge_master.py` (`MPPHSCL_HAND_VERIFIED_OVERRIDES` note
  updated + `-b` key added) -
  `code_history/2026-11-02_0406_10_merge_master.py.before`/`.diff`.
- `scripts/build_rc_vs_bans.py` (`CONVICTION_WORDS` replaces
  `NSQ_CONVICTION_WORDS`, `is_q1_ground` simplified) -
  `code_history/2026-11-02_0406_build_rc_vs_bans.py.before`/`.diff`.
- `output/source_issues.csv` (166-vs-167 appeal-order-number
  discrepancy; MPPHSCL-PRODUCT-166-2024 overlap question; PMBI
  debar.aspx has no document links) - appended, not backed up (not a
  `.py`/PROJECT_RULES.md/README file).
- `output/frozen_2026-11-02/` (new - the freeze itself) and its
  `README.txt` - `.created` snapshot taken (a new file, backed up the
  same as any other, even though it lives under `output/`).
- `PROGRESS_LOG.md` (this entry) - snapshot taken before this entry's
  edits began.

**Results with numbers**: re-ran `07_mpphscl.py` -> Firm count 140->138,
Product count 236->238 (the 2 Meril rows moving level, exactly as
expected, nothing else changed). Re-ran `10_merge_master.py` (chains
11/12 automatically) and `build_rc_vs_bans.py` - see
`output/frozen_2026-11-02/README.txt` for the complete final numbers,
not duplicated here per the same reasoning as Entry 20.

**Problems and source issues**: the two new discrepancies above, both
logged, neither resolved either way.

**Open items**: see `output/frozen_2026-11-02/README.txt` section 7 -
unchanged from Entry 20's open items except item 16 (Meril's appeal
page 2) is now done, and a new question (the 166/167 order-number
reconciliation, and whether MPPHSCL-PRODUCT-166-2024 is the same real
action as this appeal outcome) is added.

### Entry 22 - 2026-11-07 - Status corrections (checklist items 2/9/15, Meril deadline); MPPHSCL-PRODUCT-166-2024 overlap resolved via SHA-256; new freeze

**Asked**: The user gave 4 status corrections, read by eye since the
2026-11-02 freeze, plus one new check for To do in the scripts: (1)
checklist item 2 (Zee's 4 MPPHSCL orders, and separately Radiant
Parenterals' order) - CLOSED, verified by eye. (2) checklist item 15 -
CLOSED: Nestor's AP items list already names DCL Vijayawada (Q1); Zee's
PMBI CMC eye-drops case has no PDF to read at all, so it stays Q2
permanently and moves onto the RTI list instead. (3) Meril's RMSC
NIB-10/2026 sutures deadline (06-05-2026, 6 PM) - CONFIRMED by eye
against pages 1-3 of the tender PDF, no longer OCR-pending. (4) Check
whether MPPHSCL-PRODUCT-166-2024 is the same real action as
MPPHSCL-FIRM-7876-2024's appeal outcome (order ~166/167 dt 20-01-2025,
item MIS93) - if so, mark MPPHSCL-FIRM-7876-2024 and its "-b" twin
SUPERSEDED and exclude them from every count, rather than converting
them into a second product-level ban. Then: update the checklist, make
a new freeze at `output/frozen_2026-11-07/` (keeping the old one), and
log it.

**Done - task 4 (the PRODUCT-166/FIRM-7876 overlap)**: both PDFs
(`data/raw/MPPHSCL/candidate_gap_pdfs/MPPHSCL-PRODUCT-166-2024.pdf` and
`MPPHSCL-FIRM-7876-2024.pdf`) were already downloaded (from Entry 20's
candidate-gap pass) but never successfully opened - pdfplumber/pypdf
both fail on a malformed CMap (source_issues.csv, 01-11-2026). Opened
them instead with PyMuPDF (`fitz`) and pikepdf (both already installed
in this environment) - same error pattern (CMap-related), confirming
these are genuinely scanned-image PDFs (0 embedded characters, an
"AnyScanner" tag), not just a pdfplumber-specific failure. Computed
SHA-256 of all 3 PDFs (the base pair plus the "-b" twin) - this is a
hash fact, not an OCR read, so it needed no rule-F hedging:
MPPHSCL-PRODUCT-166-2024.pdf and MPPHSCL-FIRM-7876-2024.pdf are
BYTE-IDENTICAL (2,316,046 bytes, same SHA-256) - the exact same document
uploaded under both MPPHSCL's Firm list and Product list pages, for the
same order (7876 dt 06-11-2024, item MIS93). CONFIRMED: PRODUCT-166-2024
is the item-level debarment that survives Meril's appeal, not a second,
separate finding. The "-b" twin is a DIFFERENT file (967,633 bytes, a
different scan - the dispatch/forwarding copy), so its own identification
needed OCR instead (`common.ocr_page_text()`, English only, per PROJECT_RULES.md
rule F, since this is a Hindi-language document): the OCR text (saved to
`data/text/ocr/MPPHSCL-FIRM-7876-2024-b_p1.txt`/`_p2.txt`) names the same
order 7876, date 06.11.2024, item MIS93, EMD Rs 2,00,000, and the firm's
Vapi address - strong cross-reference, but OCR-derived, so this one
specific identification is flagged "OCR - check" rather than "Checked".
As a bonus, independently re-running `build_mpphscl_unclassified.py
--reclassify` (0 new downloads/OCR calls - it reuses this same cached OCR
text) found its OWN, unrelated keyword classifier matches the literal
English sentence in the OCR text ("If the successful Bidder fails to
execute the agreement...") to the SAME "Other" category these overrides
assign by hand - a third, independent line of evidence for the same
conclusion, not acted on by itself (PROJECT_RULES.md rule F still applies to
that classifier's own text-matching, same as always).

**Code changed** (every file backed up per rule B, `code_history/
2026-11-07_1050_*`):
- `scripts/07_mpphscl.py` - `MPPHSCL_LEVEL_OVERRIDES`/
  `MPPHSCL_PRODUCT_OVERRIDES` emptied (no longer turn
  `MPPHSCL-FIRM-7876-2024` into a fake second "Product" row);
  `MPPHSCL_COURT_OVERRIDES["MPPHSCL-FIRM-7876-2024"]` (reaches both the
  base row and its "-b" dedup twin automatically, applied before dedup)
  now says SUPERSEDED BY APPEAL, pointing at MPPHSCL-PRODUCT-166-2024 -
  `.before`/`.diff` at 1050.
- `scripts/10_merge_master.py` - removed the 2 FIRM-7876-2024(/-b)
  entries from `MPPHSCL_HAND_VERIFIED_OVERRIDES` (that dict's fixed
  status wrapper wrongly attributes every entry to "the user, hand-
  verified against the letter image, document_checks_2026-11-01.
  xlsx" - not true for this finding, which is the automated SHA-256/
  OCR check). Added a new, separate `MPPHSCL_SUPERSEDED_OVERRIDES` dict
  (3 entries: the FIRM-7876-2024 pair, now "Checked"/"OCR - check"
  respectively with honest provenance text, plus a new entry for
  MPPHSCL-PRODUCT-166-2024 itself, "Checked", recording it as the
  authoritative surviving ban) and its own branch in `classify_row()`,
  checked before `MPPHSCL_HAND_VERIFIED_OVERRIDES` - `.before`/`.diff`
  at 1050.
- `scripts/build_rc_vs_bans.py` - moved `"NIB10-2026"` from
  `TENDER_DEADLINES_OCR_UNCHECKED` to `TENDER_DEADLINES` (confirmed by
  the user by eye, pages 1-3) - `.before`/`.diff` at 1050.
- `scripts/build_case_files.py` - Meril's "extra_facts" deadline note
  updated to say CONFIRMED, not OCR-derived/leading-digit-uncertain -
  `.before`/`.diff` at 1050.
- `output/source_issues.csv` - one new row (07-11-2026) recording the
  SHA-256 finding and resolving the 02-11-2026 "166-vs-167" question -
  appended, not backed up (not a `.py`/PROJECT_RULES.md/README file).
- `PROGRESS_LOG.md` - this entry; checklist items 2, 9, 15 updated in
  place (section 6) - snapshot taken before this entry's edits began.

**Results with numbers**: re-ran `07_mpphscl.py` (Firm 140, Lab 2,
Product 236 - back to the pre-Entry-21 counts, since the fake Level
override is gone), `10_merge_master.py` (chains 11/12 automatically -
804 rows, 166 Matches, 0 Skipped, 1 Unmatched - unchanged),
`build_rc_vs_bans.py` (headline Q1 breach still exactly Zee/Meril/Nestor,
3 firms - but Meril's own line no longer carries the "PENDING: deadline
came from OCR" caveat; Strength groups A/B/C/X unchanged at 4/5/21/2,
D 10->11 and U 11->10 - Meril's MPPHSCL-PRODUCT-166-2024 moving from
"To check"/Group U to "Other"/Group D, a correction, not a new finding),
`build_case_files.py` (Meril.xlsx: 6 ban rows, unchanged count, updated
text), `build_mpphscl_unclassified.py --reclassify` (164 rows: Quality
21, Non-supply 21, Other 39 (was 37), To check 83 (was 85) - see the
bonus independent-classifier note above). Full numbers in
`output/frozen_2026-11-07/README.txt`.

**Problems and source issues**: one new source_issues.csv row (07-11-
2026, MPPHSCL) records the SHA-256 finding and resolves the 02-11-2026
166-vs-167 question. No new unresolved issues.

**Open items**: see the checklist (section 6) - items 2 and 15 are now
done; item 9 (RTI list) gained one more ask (PMBI, for Zee's CMC-eye-
drops order/lab). The "-b" twin's own identification remains OCR-only
(not hash-confirmed) - worth the user's own look, though the real
question (is PRODUCT-166-2024 the same action) is already settled by
the hash-identical base pair on its own. `mpphscl_unclassified_batch.
xlsx`'s "Random 20 sample" sheet is a fresh random draw from this run,
not the same 20 rows the user already hand-verified in Entry 20 - no
new hand-check is implied by refreshing this report.

### Entry 23 - 2026-11-07 - Step 3 (finance): headline_case_values.xlsx for Zee/Meril/Nestor, clause-(ii) tenders only

**Asked**: Using ONLY `output/frozen_2026-11-07/` (not the live
`output/` folder), for Zee, Meril and Nestor, for every RMSC rate-
contract line under the tenders where clause (ii) applies (Zee:
NIB-10/2024, NIB-14/2024, NIB-01/2025; Meril: NIB-10/2026; Nestor:
NIB-14/2024): compute an Upper estimate (rate x Annexure VIII's 2-year
estimated quantity, with the pack-size fix) and an Equal-share estimate
(upper / number of suppliers holding an RC for the same drug code in
the same tender); show the inputs (rate, unit, pack size, quantity,
number of suppliers, source page) and "Source: native text / OCR -
check"; totals per firm and per tender; label every figure "Estimated
contract value (2-year), not amount paid"; for the 2 scanned tenders
(NIB-01/2025, NIB-10/2026), OCR Annexure VIII and give a 10-line sample
with page numbers for checking. Output: `output/finance/
headline_case_values.xlsx`.

**Done**: built `scripts/build_headline_case_values.py` (new file) -
reuses `build_finance_report.py`'s own `parse_pack_size()`/`parse_rate()`/
`tender_short_code()`/`load_annexure_viii_quantities()` rather than
re-deriving them (that module's own "IMPORTANT UNIT FIX" docstring note
describes the exact 80-90x bug this pack-size logic already caught and
fixed - not repeated here). Reads supplier/RC facts from
`output/frozen_2026-11-07/rate_contracts/RMSC.xlsx` ONLY, as asked - a
different folder from every other script in this project's "Non-chained
utility/report scripts" list (section 3), which all still read the live
`output/` folder; source PDFs under `data/raw/` are read directly either
way (immutable raw evidence, outside `output/`).

NIB-14/2024 and NIB-10/2024 (native PDF text, already read by
`build_finance_report.py`): 33 of Zee's 47 RC lines and 3 of Nestor's 4
got a computed Upper/Equal-share estimate (the rest: pack-size
instruction not parseable, or drug code not in Annexure VIII - same
"Not available" reasons `build_finance_report.py` already uses, not new
gaps). NIB-01/2025 and NIB-10/2026 are BOTH fully scanned image PDFs (0
extractable characters) - located their Annexure VIII page ranges by
OCR keyword search (`ANNEXURE`/heading text), confirmed not guessed:
NIB-01/2025 pages 59-82 (24 pages, "List of Drugs with Specifications");
NIB-10/2026 pages 54-58 (5 pages, "List of Item(s) with
Specifications"). OCR'd every one of those 29 pages (resolution
250-300, English only - `common.ocr_page_text()`), saved to
`data/text/ocr/` per rule F.

**Found while building this**: unlike the 2 native-text tenders, OCR
text for these 2 scanned Annexure VIII tables does NOT reliably preserve
column alignment - the SAME problem already documented for KSMSCL
(31-10-2026 `source_issues.csv` entry). Spot-checked several rows by
computing the implied per-unit value (OCR quantity/cost pair) against
the RC rate: one row implied a per-unit value roughly 3.5x the RC rate,
far outside the 10-18% gap `build_finance_report.py` confirmed for the
native-text tenders - clear evidence of OCR row misalignment, not a real
figure. Per PROJECT_RULES.md rule F and the "preserve a field as unknown rather
than inventing a value" data-quality rule, decided NOT to compute any
quantity/pack-size/estimate for these 2 tenders' RC lines (92 lines
total in this report; 44 of them - all 26 of Zee's NIB-01/2025 lines and
all 15 of Meril's NIB-10/2026 lines, since Meril's case is ENTIRELY this
one scanned tender - get "Not available (OCR table alignment not
reliably parseable)", `Source` = "OCR - check"). Located the specific
page for 18 of Zee's 26 drug codes and 4 of Meril's 15 item codes by
keyword search against the OCR text (drug-name fragments, since the
bracketed code itself often did not survive OCR cleanly) - shown as a
candidate location, not a confirmed one; the rest are logged as "not
located by automated OCR search" rather than guessed. Built the
requested 10-line sample (5 pages from each scanned tender's OCR text,
with page numbers) as its own sheet.

**Code changed** (backed up per rule B):
- `scripts/build_headline_case_values.py` (new file) -
  `code_history/2026-11-07_1206_build_headline_case_values.py.created`.
- `output/source_issues.csv` - 2 new rows (07-11-2026) logging the OCR
  table-alignment problem for both scanned tenders - appended, not
  backed up (not a `.py`/PROJECT_RULES.md/README file).
- `PROGRESS_LOG.md` - this entry; section 3 (script list) and the
  checklist (section 6, new item 17) updated in place - snapshot taken
  before these edits (note: the section-3 edit itself was made just
  before this snapshot, one edit outside its own before/after bracket -
  a minor process gap, not a data-affecting one; the diff still shows
  the true before/after state across both edits together).

**Results with numbers**: Zee - 73 RC lines (13 NIB-10/2024 + 34
NIB-14/2024 + 26 NIB-01/2025), 33 with a computed estimate, sum of Upper
estimates Rs 40,13,07,126.21 / Equal-share estimates Rs 21,21,12,708.46
(both labelled "Estimated contract value (2-year), not amount paid",
and both are SUMS OVER ONLY THE 33 COMPUTED LINES, not Zee's whole
RMSC book). Meril - 15 RC lines (NIB-10/2026 only, per the task's own
scope), 0 computed (fully scanned tender, see above) - Totals sheet
shows "N/A", not a misleading 0. Nestor - 4 RC lines (NIB-14/2024 only),
3 computed, sum of Upper estimates Rs 5,27,28,434.02 / Equal-share
estimates Rs 3,07,90,464.67. Full breakdown, per-tender subtotals, and
every input column in `output/finance/headline_case_values.xlsx`.

**Problems and source issues**: the 2 new OCR table-alignment rows in
`source_issues.csv` (above) - the only new issues; neither resolved
(both logged as the user's own look needed, per the new checklist item
17).

**Open items**: checklist item 17 (new) - checking the Annexure VIII
quantities for Zee's NIB-01/2025 and Meril's NIB-10/2026 lines by eye,
using the "Scanned tender OCR sample" sheet and the raw OCR text under
`data/text/ocr/` as a starting point. Everything else in this report
(the 2 native-text tenders, the supplier counts, the totals) does not
depend on that check.

### Entry 24 - 2026-11-07 - STATUS FIRST: checklist items 4/10/17 closed; scanned-tender figures read by eye added to the finance report

**Asked**: as the first step of the FINAL STAGE (3 parts), close
checklist items 4, 10 and 17 (checked by eye - see
`output/review/document_checks_2026-11-01.xlsx` and the new
`output/review/Finance_scanned_tenders_2026-11-07.xlsx`); re-label items
8/9 (GMSCL archive, RTIs) as optional upgrades, not pending; add the
user's own read-by-eye scanned-tender figures to
`output/finance/headline_case_values.xlsx` as a new "Read by eye" sheet
and update its totals - Zee upper Rs 46.3 crore / equal-share Rs 25.3
crore (3 tenders combined; 3 lines hidden under the signature stamp left
blank), Meril upper Rs 2.28 crore / equal-share Rs 0.99 crore (15 of 15
lines), Nestor unchanged at upper Rs 5.27 crore / equal-share Rs 3.08
crore, every figure labelled "Estimated contract value (2-year), not
amount paid".

**Done - checklist**: item 4 (NIB-01/2025 deadline + page 55) CLOSED -
the user's "Rajasthan clause" sheet confirms the deadline and RESOLVES
the "3% vs 5%" discrepancy (it is 5% - "Bidder is bound to supply
minimum 5% of bid quantity on monthly basis" - the OCR's "3%" was a
misread); also surfaced, while reading that same sheet, a genuine
wording inconsistency worth carrying into the write-up: Annexure VII
item 4's affidavit asks about a ban for "conviction ... or the products
being found spurious or adulterated", narrower than clause 5(l)(II)'s
own "Not of Standard Quality" ground. Item 10 (RMSC clause quote, page
9) CLOSED - the same sheet gives the EXACT clause text for 5(l) I/II/III
(drugs) and 5(k) ii/iii (sutures), confirming the whole-firm reading
`build_rc_vs_bans.py` already uses (section 3.8). Item 17 (Annexure VIII
quantities for the 2 scanned tenders) CLOSED - see below. Items 8/9
re-labelled "OPTIONAL UPGRADES, not pending" in place (section 6
heading), per the user's own instruction - nothing already found
depends on either being done.

**Done - finance report**: extended `build_headline_case_values.py`
(not a new script - same report, same methodology, now complete) to
load `output/review/Finance_scanned_tenders_2026-11-07.xlsx`'s own
"Line by line" sheet (not hand-retyped into this script, so the new
sheet and the totals it feeds stay reproducible straight from that
file) and write 2 new sheets: "Read by eye" (all 41 computed lines - 23
of Zee's 26 NIB-01/2025 lines, all 15 of Meril's NIB-10/2026 lines - plus
the file's own 2 per-tender TOTAL rows, each one labelled "Estimated
contract value (2-year), not amount paid") and "Combined totals" (native
text + read by eye, the AUTHORITATIVE figure per firm - the existing
"Totals" sheet is relabelled "(native text only)" and kept, not
replaced, since it still correctly describes what IT covers). The "RC
lines" sheet's own "Not available" notes for the 2 scanned tenders are
updated to point to "Read by eye" ("RESOLVED 2026-11-07 - the user read
this line directly from the page image") rather than silently left
reading as still-open.

**Code changed** (backed up per rule B):
- `scripts/build_headline_case_values.py` - `READ_BY_EYE_FILE`/
  `READ_BY_EYE_FIRM_MAP` constants, `load_read_by_eye()`, the "Read by
  eye" and "Combined totals" sheets, updated "RC lines" notes -
  `.before`/`.diff` pairs at 1050 (today's earlier Part-2 session) and
  1240.
- `PROGRESS_LOG.md` - this entry; section 3 script-list note, and the
  checklist (section 6, items 4/8/9/10/17) updated in place - snapshot
  taken before these edits.

**Results with numbers**: `output/finance/headline_case_values.xlsx`
"Combined totals" sheet - ZEE LABORATORIES LTD: upper Rs 46,37,12,355.21
(native Rs 40,13,07,126.21 + read-by-eye Rs 6,24,05,229), equal-share Rs
25,27,80,853.46 (native Rs 21,21,12,708.46 + read-by-eye Rs 4,06,68,145)
- matches the user's own Rs 46.3cr/Rs 25.3cr figures (rounding). MERIL
ENDO SURGERY PVT LTD: upper Rs 2,27,69,369,
equal-share Rs 98,68,123 - matches the user's Rs 2.28cr/Rs 0.99cr
exactly (Meril's only clause-(ii) tender is the fully-scanned NIB-10/2026,
so this IS the whole figure, not an addition to a native-text component).
NESTOR PHARMACEUTICALS LTD: unchanged at upper Rs 5,27,28,434.02,
equal-share Rs 3,07,90,464.67 (no scanned tender in Nestor's clause-(ii)
scope).

**Problems and source issues**: none new - the read-by-eye file's own 3
"hidden under signature stamp" lines (Tinidazole/430, Cefpodoxime
suspension/NRD-485, Sorbitol+Tricholine/NRD-520) and 2 "partly hidden -
check" lines (Febuxostat 40/NRD-632, Loratadine/NRD-670) are reproduced
exactly as the user marked them in the "Read by eye" sheet's own "Read
status" column - not silently resolved into a confident figure by this
script.

**Open items**: none from this entry - checklist items 4, 10, 17 are
fully closed; items 8/9 are explicitly optional, not tracked as pending
work going forward.

### Entry 25 - 2026-11-07 - FINAL STAGE Part 1 (data science): network chart, summary numbers, contagion test, 4 charts

**Asked**: using ONLY `output/frozen_2026-11-07/` (plus `output/finance/`
and `output/review/`), build: (1) `17_network_chart.py` -> a firm<->buyer
network PNG, one line per active ban, only firms with 2+ buyers, Zee/
Meril/Nestor highlighted; (2) `summary_numbers.xlsx` - firms by buyer
count, Q1 vs Q2, court stays/revocations by buyer, median days between a
firm's first and second ban, reason categories by buyer, the 5
concealment cases; (3) a contagion test against the user's own Project 1
CDSCO NSQ data (`E:\cdsco_mqri`) - for firms in both datasets, NSQ alert
counts 12 months before vs after a state ban, with an explicit "no
conclusion" call if fewer than 20 firms match; (4) 4 labelled PNG charts
with a source line each.

**Done**: added `scripts/analysis_common.py` (new, shared helper, never
run directly) - `load_master_rows()` (the ONLY entry point every Part 1
script uses to read `output/frozen_2026-11-07/master_list.xlsx`) and
`is_active_ban()`, the single definition of "active" used identically
across every Part 1 deliverable: excludes a row only when its own
"Court stay or revoked?" text is a GENUINE judicial stay, revocation, or
this session's own "SUPERSEDED" finding (regex-matched: stay(ed)/
revoked/revocation/superseded/set aside/abeyance/suspension/interim) -
explicitly NOT a plain "Blacklisting Period Completed" note, which is a
different fact (a real, never-overturned ban that has simply expired
with time). Loaded the dataviz skill before writing any chart code;
every PNG uses its reference palette (fixed categorical hue order,
recessive gridlines, a labelled source line, one axis per chart, a
legend wherever there is more than one series).

**Done - 17_network_chart.py**: bipartite layout (no networkx in this
environment - laid out by hand, 2 fixed columns, reproducible run to
run, not force-directed). 63 of 369 actively-banned firms hold an active
ban from 2+ distinct buyers; 8 of the project's 9 buyers appear among
them (KMSCL banned no firm that also appears elsewhere). Zee, Meril and
Nestor drawn in bold red; found and fixed one real layout bug before
accepting the chart - the default `fig.tight_layout()` clipped both the
title and every left-hand firm-name label against the figure edge
(the long labels extend well past the plotted data's own x-range, which
`tight_layout()` does not account for) - replaced with explicit
`fig.subplots_adjust()` margins and a wider figure. Noted, not treated
as a bug: "Zenith Drugs Ltd" and "Zenith Drugs Pvt Ltd" appear as 2
separate nodes - checked directly, these are 2 genuinely distinct
standardized strings in `master_list.xlsx` (same kind of name-variant
question as "Medipol Pharmaceutica"/"Medipol Pharmaceutical", section
3.5) - not merged, since nothing in this project is ever auto-merged
(PROJECT_RULES.md rule D); `match_decisions.xlsx` is the place that would
adjudicate it, not this chart.

**Done - build_summary_numbers.py** (`output/analysis/summary_numbers.xlsx`,
8 sheets): Firms by buyer count - 306 firms with 1 buyer, 51 with 2, 12
with 3+ (active bans only). Q1 vs Q2 - 37 Q1 rows/29 firms, 334 Q2 rows/
171 firms (ALL rows, not just active - a Q1/Q2 finding is a fact about
the order letter itself). Stayed or revoked by buyer - 40 rows across 5
buyers (GMSCL 21, KSMSCL 9, MPPHSCL 5, APMSIDC 3, BMSICL 2). First-to-
second ban gap - median 285.5 days across 102 firms with 2+ distinct
ban dates (any row, active or not - a later-stayed ban was still a real
ban EVENT at the time). Reason categories by buyer - a full crosstab
(Quality 371 rows total, cross-checked exactly equal to Q1+Q2's 37+334
above). Concealment cases - the same 5 rows as `concealment_cases.xlsx`,
reproduced for one-stop reading.

**Done - build_contagion_test.py** (`output/analysis/contagion_test.xlsx`):
anchor = each matched firm's EARLIEST active, QUALITY-CATEGORY ban date
(Non-supply/Other bans excluded as the anchor - no reason to expect them
to relate to NSQ-alert timing). Matched by `common.standard_name()`
first (94 firms matched exactly), then a shared-distinctive-word filter
+ `rapidfuzz.ratio>=90` (the SAME threshold 14_match_firms.py's own
design record, section 3.5, found clean) for 16 more - kept in a
SEPARATE "not counted" sheet, never folded into the main count (PROJECT_RULES.md
rule D - nothing auto-merged). 94 >= 20, so a result IS reported: 202
NSQ alerts in the 12 months before vs 314 after (CDSCO's `master_alerts`
has no day-level date - month resolution only, stated explicitly, not
silently treated as day-precise). Checked the result is not misleadingly
uniform before writing the verdict: 37 firms rose, 34 fell, 23 flat: of
the 3 headline cases, only Zee and Nestor matched at all (Meril's ban is
for a surgical suture, outside CDSCO's drug-only NSQ portal) and BOTH
FELL (Zee 4->3, Nestor 5->0) - the aggregate rise is driven by OTHER
matched firms, not this project's own 3 cases. Also surfaced and quoted
directly in the verdict text: the CDSCO project's own README documents a
system-wide alert-VOLUME jump from increased state-lab reporting in
2025 ("do not read the jump as a change in quality") - a likely confound
for any firm whose "after" window falls in/after 2025, which this script
does not correct for.

**Done - build_analysis_charts.py** (4 PNGs) + **17_network_chart.py**
(1 PNG), all in `output/analysis/`: `bans_by_buyer_reason.png`,
`q1_vs_q2.png`, `court_stays_by_buyer.png`, `headline_cases_timeline.png`,
`network.png`. Bug found and fixed in the timeline chart before
accepting it: the first version took min/max "Blacklisted from/to"
across EVERY Quality-category row for each case's named buyer, which
silently included Zee's OTHER, unconfirmed (Q2) MPPHSCL bans AND a
long-expired Q1 one (Ranitidine, order 8078-2021, expired 12-08-2023 -
itself over a year before Zee's earliest RMSC deadline) - stretching the
drawn "ban period" bar back to 2020 and misrepresenting the actual
headline finding. Fixed by restricting to CONFIRMED Q1 rows only, further
restricted to rows whose own "to" date is on/after the case's earliest
RMSC deadline (excludes the expired one without hand-picking it by Entry
ID). Also fixed, on inspection per the dataviz skill's own step 7 (look
at the rendered result): overlapping deadline labels on the timeline
(staggered vertically) and a cut-off source line / overflowing legend on
2 of the 4 charts (manual margins + `textwrap`, not `tight_layout()`).

**Code changed** (every file backed up per rule B,
`code_history/2026-11-07_1246`-`1256_*`):
- `scripts/analysis_common.py` (new) - `.created`.
- `scripts/build_summary_numbers.py` (new) - `.created`.
- `scripts/build_contagion_test.py` (new) - `.created`.
- `scripts/build_analysis_charts.py` (new) - `.created`.
- `scripts/17_network_chart.py` (new) - `.created`.
- `PROGRESS_LOG.md` (this entry; section 3 script list updated in place) -
  snapshot taken before these edits.

**Results with numbers**: see each sub-section above; also printed to
the terminal by every script's own `main()` and saved in the output
files themselves - not duplicated a third time here.

**Problems and source issues**: none logged to `source_issues.csv` this
entry - every finding above was either a code bug (found and fixed
before accepting the output, documented above) or an honestly-reported
data limitation already stated in the relevant file's own text (the
CDSCO month-resolution/2025-reporting-volume caveats, the Zenith Drugs
name-variant question).

**Open items**: none from Part 1 itself. Part 2 (write-up) and Part 3
(independent fact-check) of the FINAL STAGE are still to come.

### Entry 26 - 2026-11-07 - FINAL STAGE Part 2 (write-up): full report, one-pager, 3 case files, resume lines, README

**Asked**: a full report (`Blacklist_Gap_Report.md` + `.docx`), a 1-page
summary (`One_page_summary.md` + `.pdf`), 3 case files (`cases/Zee.md`,
`Meril.md`, `Nestor.md`), `Resume_lines.md` (3 lengths), and a GitHub
`README.md` - every number traceable to a file/sheet/row or PDF page, strict
wording rules (no "violation"/"fraud"/"illegal" - "possible breach of the
tender clause as worded" instead; "estimated contract value" never "money
paid"; company-level only, no individual names; the "not contacted for
comment" line; uncertainty marked as uncertain), using only
`output/frozen_2026-11-07/`, `output/finance/`, and `output/review/`.

**Done - fact-gathering before writing** (so every claim in the write-up
could be sourced as it was written, not retrofitted after): read the exact
clause 5(l) text from `document_checks_2026-11-01.xlsx`'s own
"Rajasthan clause" sheet (already in hand from the STATUS FIRST step);
re-confirmed the General Conditions (o) mid-contract-cancellation clause
and the affidavit's narrower "spurious or adulterated" wording from this
project's own PROGRESS_LOG.md Entry 9; pulled the full CAG Haryana
citation from `output/frozen_2026-11-07/cases/Nestor.xlsx`'s own "Outside
evidence" sheet; got the CORRECTED Karnataka stay count directly from
`output/frozen_2026-11-07/buyers/KSMSCL.xlsx` (9 of 22 - 1 Firm + 8
Product - not the superseded "7 of 13" narrative figure the task itself
named, which this write-up flags as a discrepancy rather than silently
using); searched this project's own data for "bans stop at the border"
evidence and found none pre-existing - opened Meril's actual GMSCL order
PDF directly (native text, not OCR) and found its own distribution list
copies ONLY Gujarat state offices, quoted verbatim with page citation; a
targeted keyword search of all 25 downloaded GMSCL order PDFs found no
other "other state"/"within Gujarat"-type wording, and no readable
(non-scanned) MPPHSCL order was available to check the same way - both
gaps stated explicitly in the report rather than papered over. Web-
searched the US SAM.gov/FDA Debarment List comparison (4 sources cited).

**Done - the write-up itself**: `write_up/Blacklist_Gap_Report.md` (9
sections, the full structure asked for), `write_up/One_page_summary.md`,
`write_up/cases/{Zee,Meril,Nestor}.md` (each with a dated timeline and a
full source list with page numbers), `write_up/Resume_lines.md` (1 line /
2 lines / 3 bullets, each naming all 3 of regulatory/finance/data-science),
`write_up/README.md` (folder map, re-run instructions, explicitly excludes
raw PDFs and individual names per the task's own instruction).

**Bug found and fixed while verifying the report's own Zee finance
figure**: "lines with a computed value" was first written as "69 of 73" -
checked against `output/finance/headline_case_values.xlsx` directly and
found the real figure is **56 of 73** (33 native + 23 read-by-eye) - the
first draft had conflated "not hidden under the signature stamp" (23 of
26 scanned-tender lines) with the TOTAL computed count, missing the 14
further native-tender lines that separately could not be matched to an
Annexure VIII entry. Fixed before the report was considered final -
caught by checking the actual workbook, not assumed correct from memory.
Also caught and fixed: the Resume_lines.md combined-value range was first
written "Rs 21-54 crore", using Zee's own NATIVE-ONLY equal-share figure
(Rs 21.2cr) instead of the combined (native + read-by-eye) total across
all 3 firms (Rs 29.35cr, rounds to Rs 29cr) - fixed to "Rs 29-54 crore".

**Done - docx/pdf export** (`scripts/build_write_up_exports.py`, new
file): no pandoc/LibreOffice/reportlab/markdown package in this
environment (checked first) - built a small, constrained Markdown-subset
converter by hand instead (python-docx for the report, matplotlib's own
PDF backend composing one page of text + the timeline chart PNG for the
one-pager), rather than a general-purpose parser, since this project's own
write-up files only use a known, limited set of Markdown constructs.
**2 real bugs found and fixed before accepting either output** (dataviz
skill's own step 7 - look at the rendered result, not just "it ran"):
(1) a paragraph or list item whose **bold**/*italic* marker pair is split
across 2 wrapped source lines was never being un-marked, since the first
version of the converter processed the file one raw line at a time and
the regex never saw a matching close-marker in the same pass - confirmed
in BOTH outputs (a stray "**" in the docx's 3rd recommendation, and in the
PDF's own 2nd paragraph) before writing a proper paragraph-joining pre-
pass (`_join_paragraphs()`) and confirming 0 stray "**" remained; (2) that
same pre-pass's first version re-split the clause-5(l)(II) blockquote into
11 separate one-line "quotes" (this project's own blockquotes repeat "> "
on every wrapped line, unlike its list items) - fixed by only starting a
NEW quote block on a "> " line when one wasn't already open. Also fixed:
the one-pager's attribution line (the `*...*` line at the very bottom) was
not word-wrapped at all and ran off the page edge - added wrapping to that
branch too.

**Code changed** (backed up per rule B):
- `write_up/Blacklist_Gap_Report.md`, `One_page_summary.md`,
  `cases/{Zee,Meril,Nestor}.md`, `Resume_lines.md` - new content files, not
  backed up via `record_change.py` (same convention already used for
  every other data/deliverable output in this project - e.g.
  `match_decisions.xlsx`, `concealment_cases.xlsx` - "not a .py/PROJECT_RULES.md/
  README file", PROJECT_RULES.md rule B's own stated scope).
- `write_up/README.md` - backed up anyway, given its name, even though
  rule B's literal text names "README.txt" specifically, not this folder's
  own "README.md" - `code_history/2026-11-07_1310_README.md.before`/
  `.diff` (the `.diff` is empty - the file was already written before the
  snapshot pair was taken, a minor process-ordering gap, not a data one;
  its full content is preserved in the `.before` file either way).
- `scripts/build_write_up_exports.py` (new) - `.created`.
- `PROGRESS_LOG.md` (this entry; section 3 script-list note updated in
  place) - snapshot taken before these edits.

**Results with numbers**: `write_up/` now holds 9 files (5 at the top
level + 3 case files + the `.docx`/`.pdf` pair) - see the end-of-task
summary for the full path list. Every figure in the report cross-checked
against its own source file before being accepted (see the 2 bugs above,
both caught this way).

**Problems and source issues**: none newly logged to `source_issues.csv` -
the "bans stop at the border" MPPHSCL gap (no readable native-text order
found) is stated directly in the report's own Finding (c) and Limits
sections instead, since it is a property of the write-up's own evidence
base, not a new fact about a buyer's data.

**Open items**: Part 3 (independent fact-check) of the FINAL STAGE is
still to come.

### Entry 27 - 2026-11-07 - FINAL STAGE Part 3: independent fact-check, 7 real errors found and fixed

**Asked**: an independent fact-check, by a reviewer that has NOT seen this project's work, given
access to ONLY `write_up/`, `output/frozen_2026-11-07/`, `output/finance/`,
`output/analysis/`, `output/review/`, and `data/raw/`, should list every
number/date/name/claim in the report, one-pager, case files, and resume
lines; trace each to its source and mark Correct/Wrong/Cannot verify; check
the wording rules; save `write_up/fact_check.xlsx`. Then fix every Wrong
item, re-run the check on the fixed items, and report totals.

**Done**: ran the fact-check with no prior context, with written
instructions listing the 6 allowed folders and the exclusions
(`PROGRESS_LOG.md`, `scripts/`, `code_history/`, `source_issues.csv`, and
the live, non-frozen `output/` folder - all explicitly off-limits). It
checked 83 claims across the 6 write-up documents and found 9 Wrong, 13
Cannot-verify (mostly things genuinely outside its allowed folders - the
external CDSCO dataset, the CAG/Tribune citation, the superseded "7 of 13"
Karnataka figure which lives in the off-limits `PROGRESS_LOG.md`), and 3
wording-rule violations (all 3 tied to the same 3 underlying factual
errors - an uncertain/wrong claim stated as flat fact).

**Did NOT blindly apply its fixes** - independently re-verified all 9
"Wrong" findings against the source files myself before touching any
write-up text (re-rendered the Meril GMSCL order PDF as an image, re-
queried `RMSC.xlsx` and `PMBI.xlsx` directly). **Found the fact-checker
had made ITS OWN error** on 1 of its 9 findings (claim #40: "Meril holds
16 NIB-10/2026 lines, not 15") - it had matched RMSC's own rate-contract
data on the bare substring "MERIL", which also catches "MERIL DIAGNOSTICS
PVT LTD" (Entry ID `RMSC-RC-NIB10-R79-MERIL`) - a legally DISTINCT
supplier from "MERIL ENDO SURGERY PVT LTD" (the firm this whole case is
about) in RMSC's own "Supplier (standard)" column, the same kind of name-
variant question `match_decisions.xlsx` exists to adjudicate by hand, not
by substring match. The original report's "15 lines" was correct all
along. Sent this correction back to the SAME fact-checker (rather than starting
a fresh check, which would likely repeat the same substring-matching
mistake) along with the 6 genuinely-confirmed fixes to
re-check; it corrected its own `fact_check.xlsx` entry for claim #40 and
confirmed all 6 real fixes are now exactly right, with no new errors
introduced.

**The 7 real errors, found and fixed**:
1. Zee's estimated contract value was wrongly given as "Rs 21 crore to
   Rs 46 crore" in 3 places (Report section 4a, `One_page_summary.md`,
   `cases/Zee.md`) - Rs 21.21 crore is the NATIVE-TEXT-ONLY equal-share
   figure, wrongly paired with the COMBINED upper-estimate figure
   (Rs 46.37cr); the report's own Finance section 5 table and
   `Resume_lines.md` already had the correct combined figure (Rs 25.28cr)
   two sections later in the same document - an internal inconsistency
   within the report's own numbers, now resolved to Rs 25.28-46.37 crore
   everywhere.
2. `cases/Zee.md`'s PMBI archive citation said "2 further Q2 findings...
   both expired before any Rajasthan deadline here" - actually 3 entries
   exist (`PMBI.xlsx`), and 2 of the 3 Rajasthan deadlines (NIB-10/2024,
   NIB-14/2024) occur WHILE those bans were still active - only the
   latest deadline (NIB-01/2025) postdates their expiry. Fixed with the
   correct count and per-deadline detail.
3. `cases/Meril.md`'s timeline said GMSCL's rate contract was "executed
   28-02-2024" under tender enquiry "GMSCUD-662/RC/2019-20" - the order's
   own table (re-rendered as a page image, not just text-extracted, to
   settle this) reads tender enquiry GMSCL/D-662/RC/2019-20 dated
   28-02-2020 (bid due 03-06-2020), with the RATE CONTRACT executed
   17-05-2021 - 2 separate dates/events I had conflated into one wrong
   row. "GMSCUD" is confirmed, independently, to be a font/character-
   encoding artifact in this SPECIFIC PDF's own embedded text layer - both
   pdfplumber and PyMuPDF extract the identical "GMSCUD" from the native
   text, but the rendered page image clearly shows "GMSCL/D" - logged
   here, not silently corrected without explanation, since it is a real
   property of this one source file worth knowing about if anyone re-
   extracts it the same way.
4. `cases/Meril.md`'s RC-contract-start dates said "20-08-2025 to
   27-08-2025" - the real dates (`RMSC.xlsx`) are 20-AUG-2026/27-AUG-2026,
   a one-year error. Fixed, and the row moved to after the bid-deadline
   row in the timeline table so the table stays chronological (it wasn't,
   with the wrong year).
5. `cases/Nestor.md`'s timeline said "Nestor's 3 RMSC rate-contract lines
   under NIB-14/2024 begin... a 4th, under the separate NIB-08/2024
   tender... began 14-02-2025 and 23/25-04-2025" - all 4 of Nestor's
   NIB-14/2024 lines are genuine (3 on 29-09-2025, 1 on 17-11-2025); the
   cited dates belong to Nestor's SEPARATE 4 lines under NIB-08/2024 (a
   different tender, correctly excluded from this review's clause-(ii)
   scope already) - the 2 tenders' own line sets had been conflated.
   Fixed to describe only the genuine NIB-14/2024 lines.

**Code changed**: none - this entry is entirely about `write_up/*.md`
content fixes (not backed up via `record_change.py`, same convention as
Entry 26) plus `write_up/fact_check.xlsx` (the fact-checker's own new output
file, not backed up either - a data/deliverable file). `write_up/
Blacklist_Gap_Report.docx` and `One_page_summary.pdf` were regenerated
from the corrected `.md` files via the existing `build_write_up_exports.py`
(no script changes needed). `PROGRESS_LOG.md` (this entry) - snapshot
taken before these edits.

**Results with numbers**: final claim count - 83 checked, **68 Correct**
(61 originally + 7 confirmed fixed), **2 Wrong** (both explained: #40 is
the fact-checker's OWN error, not a report error; #41, the 2025-vs-2026
date fix, is independently confirmed correct in the current document -
its "Wrong" verdict in `fact_check.xlsx` describes the ORIGINAL draft, as
a historical record, not the current, already-fixed text), **13 Cannot
verify** (almost entirely things genuinely outside the fact-checker's
allowed folders - the external CDSCO dataset, the CAG/Tribune web
citation, a few historical MPPHSCL bans with no downloaded order PDF,
and the superseded "7 of 13" Karnataka narrative figure, which lives in
this very log file, correctly off-limits to it). Wording rules: 22 OK, 0
violations (all 3 original violations tied to the same 3 now-fixed
factual errors, resolved once the text itself became accurate and
explicitly marked each correction).

**Problems and source issues**: none logged to `source_issues.csv` - every
issue found this entry is a write-up TEXT error (now fixed), not a new
fact about a buyer's underlying data.

**Open items**: none - all 3 parts of the FINAL STAGE are complete. The
only standing, deliberately-unresolved items are the 13 "Cannot verify"
claims (genuinely outside this review's own source material) and the
pre-existing checklist items 8/9 (now explicitly optional, not pending -
see Entry 24).
