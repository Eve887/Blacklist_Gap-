BLACKLIST GAP - DATA FREEZE - 2026-11-07
=========================================

This folder is a complete, point-in-time copy of everything in output\
as of 2026-11-07, after applying the user's status corrections (checklist
items 2, 9, 15; the Meril NIB-10/2026 sutures deadline) and an automated
SHA-256/OCR check resolving the MPPHSCL-FIRM-7876-2024 /
MPPHSCL-PRODUCT-166-2024 overlap flagged at the previous freeze.

FROM NOW ON, ALL WRITE-UP NUMBERS COME FROM THIS FOLDER (including the
finance report at output/finance/headline_case_values.xlsx), not from the
live output\ folder or from output/frozen_2026-11-02/, both of which
predate this freeze's corrections.

See PROGRESS_LOG.md Entry 22 for the full step-by-step record of how this
was built; Entry 20/21 and output/frozen_2026-11-02/README.txt for
everything carried forward unchanged from the previous freeze.


1. ROW COUNTS
--------------

master_list.xlsx: 804 rows total (unchanged from the 2026-11-02 freeze -
  this freeze corrects classification/status text and two tender
  deadlines; it adds or removes no rows).
  Per buyer: APMSIDC 43, BMSICL 81, GMSCL 99, KMSCL 5, KSMSCL 22,
             MPPHSCL 378, PMBI 75, RMSC 101.
  Matches sheet: 166. Skipped rows: 0. Unmatched revocations: 1.

mpphscl_unclassified_batch.xlsx: 164 rows - Quality 21, Non-supply 21,
  Other 39 (was 37), To check 83 (was 85) - refreshed from cache (0 new
  downloads, 0 new OCR calls) after this freeze's 3 newly-OCR'd Meril
  PDFs (MPPHSCL-PRODUCT-166-2024, MPPHSCL-FIRM-7876-2024, its "-b" twin)
  gave the existing classifier's "agreement not executed" signal
  something to match for the first time. This is a secondary utility
  report, not master_list.xlsx itself - master_list.xlsx already carried
  the correct categories for these 3 rows via 10_merge_master.py's own
  overrides (see section 4 below), independent of this refresh.

Everything else in section 1 of the 2026-11-02 freeze's README
(cross_buyer_cases.xlsx, candidate_gaps.xlsx, match_candidates.xlsx,
match_decisions.xlsx, concealment_cases.xlsx) is UNCHANGED - none of
those scripts were touched this freeze.


2. GROUPS (rc_vs_bans.xlsx "Strength" sheet - distinct firms per group,
   a firm can appear in more than one)
------------------------------------------------------------------------

  A (firm-level ban, Quality):                4 firms  (unchanged)
  B (product-level, same drug+dosage form):    5 firms  (unchanged)
  C (product-level, different drug/form):     21 firms  (unchanged)
  D (non-quality reason - Non-supply/Other):  11 firms  (was 10)
  U (reason genuinely unknown - "To check"):  10 firms  (was 11)
  X (stayed or revoked):                       2 firms  (unchanged)

The D/U shift (1 firm) is Meril: MPPHSCL-PRODUCT-166-2024 moved from
category "To check" (Group U) to "Other" (Group D) once it was given a
real, verified ground this freeze (see section 4) - a correction, not a
new finding; Meril's own A/B/C headline membership (from its GMSCL ban)
is untouched by this.


3. Q1/Q2 COUNTS
-----------------

Master-list-wide (every buyer, every row) - UNCHANGED from the 2026-11-02
freeze (today's fixes touched Category "Other" rows and two tender
deadlines, nothing Quality-category):
  Q1 (named statutory lab): 37 rows, 29 distinct firms.
  Q2 (buyer's own/empanelled lab, or no lab named): 334 rows, 171
    distinct firms (161 of those have NO Q1 row anywhere).
  N/A (not a Quality-category row): 433 rows.

RMSC-clause breach check (rc_vs_bans.xlsx, Groups A/B/C only, clause
5(l)(II)/5(k)(ii) - whole firm):
  Q1 HEADLINE breach: 3 firms (section 4 below) - UNCHANGED membership,
    but Meril's own entry no longer carries the "PENDING - OCR deadline"
    caveat (see below).
  Q2-conditional breach ("if Q2 counts as NSQ ground"): 10 firms
    (section 5 below) - unchanged.
  Policy-gap firms (Q1 ground exists, not active at the bid deadline):
    0 firms - unchanged.
  Product-only reading (clauses 5(l)(I)/(III), 5(k)(iii) - Group B
    only): 0 firms - unchanged.


4. THE 3 HEADLINE CASES (Q1, clause ii - whole firm)
------------------------------------------------------

1. ZEE LABORATORIES LTD - Madhya Pradesh (MPPHSCL) banned 3 Zee
   injections as NSQ on Central Drugs Laboratory, Kolkata reports (June/
   July/Nov 2024). Active on 3 Rajasthan tender deadlines (NIB-10/2024,
   NIB-14/2024, NIB-01/2025). Zee holds 73 RMSC contract lines.
   CLOSED 2026-11-07 (checklist item 2): the user confirmed all 4 MPPHSCL
   order PDFs by eye - no change to the data, substance was already
   correct.
   Separately, Zee's PMBI debarment for Carboxy Methyl Cellulose Eye
   Drops (18-12-2024) stays Q2 PERMANENTLY, not "until the PMBI order is
   read" (checklist item 15, CLOSED 2026-11-07): no order PDF exists for
   this row at all - PMBI's debar.aspx exposes no document link for ANY
   row (confirmed 02-11-2026). This has been added to the RTI list
   (checklist item 9) as the only remaining way to find the lab. Still
   NOT active on the relevant bid deadline - the POLICY-GAP-SHAPED
   finding, reported under section 5 (Q2-conditional), not as a Q1
   policy gap.

2. MERIL ENDO SURGERY PVT LTD - Gujarat (GMSCL) debarred its silk suture
   (item-level) after the Food and Drug Laboratory, Vadodara declared it
   NSQ (sterility). Active on the RMSC NIB-10/2026 sutures tender
   deadline. STATUS: CONFIRMED, not pending, as of this freeze - the
   user checked the deadline (06-05-2026, 6 PM) by eye against pages 1-3
   of the tender PDF, 2026-11-07 (PROJECT_RULES.md rule F satisfied; previously
   OCR-unchecked). Meril holds 17 RMSC RC lines (sutures).
   Separately (RESOLVED 2026-11-07, following up the 166-vs-167 question
   left open at the 2026-11-02 freeze): MPPHSCL-PRODUCT-166-2024 is
   CONFIRMED the SAME real action as Meril's MPPHSCL appeal outcome
   (order 7876 dt 06-11-2024, item MIS93) - MPPHSCL-PRODUCT-166-2024.pdf
   and MPPHSCL-FIRM-7876-2024.pdf are SHA-256-IDENTICAL files (a hash
   fact, not an OCR read). MPPHSCL-FIRM-7876-2024 and its "-b" dedup twin
   are now marked SUPERSEDED BY APPEAL and excluded from every active-ban
   count (their own Court-stay field, which build_rc_vs_bans.py's
   classify() always treats as group "X"); MPPHSCL-PRODUCT-166-2024 is
   the single, authoritative row for the surviving item-level MIS93
   debarment - a rate-quoting/agreement-execution error (Category
   "Other"), NOT a quality finding, and does not affect Meril's GMSCL-
   based headline case above, which is a separate ban by a separate
   buyer. The "-b" twin's own identification as the same action rests on
   an OCR cross-reference (not a hash match - its PDF is a different
   scan), so it is marked "OCR - check" pending the user's own look, even
   though the overall conclusion is corroborated three ways (hash match
   on the base pair, OCR text cross-reference on "-b", and an independent
   keyword-classifier match on the same OCR text - see
   output/source_issues.csv, 07-11-2026).

3. NESTOR PHARMACEUTICALS LTD - Andhra Pradesh (APMSIDC) banned its iron
   + folic acid syrup on a DCL Vijayawada report. Active on the RMSC
   NIB-14/2024 deadline. Nestor holds 4 RMSC contract lines under that
   tender (+4 more under NIB-08/2024, whose own deadline is unknown -
   tender PDF not found).
   CLOSED 2026-11-07 (checklist item 15): AP's own items list already
   names "DCL, Vijayawada for 4 Batches" in its "Lab that declared NSQ"
   field (AP-P-2023-24-005) - Q1, confirmed, no code or data change was
   needed; this was already correct in master_list.xlsx before this
   freeze.

NOTE (carried forward from 2026-11-02, unchanged): Protech Telelinks
(APMSIDC) is NOT a headline case - AP's own "NSQ declared by" field
names no lab (Q2). Sits in section 5 below.


5. SECOND-TIER LIST (Q2-conditional breach - "possible breach if Q2
   counts as NSQ ground") - UNCHANGED from the 2026-11-02 freeze
-----------------------------------------------------------------------

Bharat Parenterals Ltd, Divine Laboratories Pvt Ltd, Healers Lab, Modern
Laboratories, Nandani Medical Laboratories Pvt Ltd, Omega Pharma,
Protech Telelinks, Vivek Pharmachem India Ltd, Zee Laboratories Ltd
(the PMBI CMC-eye-drops case specifically - see section 4), Zenith
Drugs Pvt Ltd.

These are quality findings where NO statutory lab is named (the buyer's
own/empanelled lab, or no lab at all) - genuine quality signals "in
substance" per the user's own standard, but not a formal Section-25
NSQ declaration. Reported separately by design; never counted toward
the Q1 headline.


6. CONCEALMENT CASES (output/concealment_cases.xlsx) - UNCHANGED from
   the 2026-11-02 freeze, now additionally eye-verified for Radiant
--------------------------------------------------------------------

1. Unicure India Ltd - caught by GMSCL (Gujarat) - "suppression/non
   disclosure of debarment" (26-06-2025 to 25-06-2027).
2. Divine Laboratories Pvt Ltd - caught by GMSCL - "suppression/non
   disclosure of debarment" (04-02-2026 to 03-02-2027).
3. Shiv Paramount Organic And Fertilizers Pvt Ltd - caught by GMSCL -
   "Submission of Wrong/false/misleading Self-Attested Declaration of
   Debarment" (31-07-2026 to 30-07-2031).
4. Medipol Pharmaceuticals India Pvt Ltd - caught by PMBI - PMBI's OWN
   revocation of its earlier action names "Concealing the fact about
   blacklisting of the firm by GMSD" (a different procurement
   authority). Not a normal master_list.xlsx row - see
   output/source_issues.csv and buyers/PMBI.xlsx "Unmatched
   revocations".
5. Radiant Parenterals Ltd - caught by MPPHSCL (Madhya Pradesh) - did
   not declare that TNMSC (Tamil Nadu) had blacklisted its Phenytoin
   Sodium Injection on quality grounds; false affidavit; debarred 5
   years (order 273 dt 27-01-2025). CLOSED 2026-11-07 (checklist item
   2): the user also confirmed this order PDF by eye - no data change,
   this row's category/status were already "Checked" since Entry 20.


7. OPEN ITEMS AT THE TIME OF THIS FREEZE
--------------------------------------------

Confirmation still needed (the user's own checklist, PROGRESS_LOG.md
section 6):
  - Save the GMSCL Archive (Blacklisted) tab by hand - blocked by
    CAPTCHA for automated capture (item 8).
  - File RTIs, including the Zee/Meril affidavit questions AND (new
    this freeze) PMBI - for Zee's CMC-eye-drops order or its lab (item 9).
  - The "-b" twin's identification as part of Meril's MPPHSCL
    supersession (section 4) is OCR-derived (PROJECT_RULES.md rule F) - still
    worth the user's own look at MPPHSCL-FIRM-7876-2024-b.pdf, though
    the base-pair hash match already settles the real question (is
    PRODUCT-166-2024 the same action) on its own.
  - mpphscl_unclassified_batch.xlsx's "Random 20 sample" sheet is a
    FRESH random draw from this freeze's refreshed 164-row batch, not
    the same 20 rows the user already hand-verified in Entry 20 - no
    new hand-check is implied or requested by this freeze.

Logged discrepancies, not resolved either way (output/source_issues.csv) -
unchanged from the 2026-11-02 freeze EXCEPT the 166/167 order-number
item, now resolved (see below):
  - ANG's own order letter is numbered 3099; this project's Entry ID
    (from RMSC's live page) uses 3078.
  - Karnataka's "Read me" narrative says 21 rows/13 products/7 stayed;
    the detailed sheet itself has 22 rows/14 products/9 stayed.
  - PMBI's debar.aspx exposes no document link for ANY row (confirmed
    by inspecting the live page's HTML directly) - Zee's CMC-eye-drops
    PDF does not exist via this page; nothing was downloaded. Now also
    on the RTI list (item 9).
  - RESOLVED 2026-11-07: the 166-vs-167 appeal-order-number question and
    whether MPPHSCL-PRODUCT-166-2024 is the same real action as the
    appeal outcome - YES, confirmed via SHA-256 (see section 4). "166"
    (Product list, SR/166/F791/2024) and "7876" (Firm list, 7876/F605/
    2024) are MPPHSCL's own two different internal reference numbers
    for the SAME order on its two separate list pages, not a data error
    either way.

Not independently verified by this project (confirmed real by the
user's own document check, but still only as strong as that read):
  - All Q1 lab-naming facts (Zee x4, RMSC letters x6) came from the
    user reading the order PDFs directly, not from automated
    extraction.
  - The whole-firm clause reading itself (section 2-5 above) depends on
    the exact clause text quoted in document_checks_2026-11-01.xlsx
    "Rajasthan clause" sheet, read by the user from the tender PDF pages.
