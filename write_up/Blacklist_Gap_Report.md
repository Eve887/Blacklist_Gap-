# The Blacklist Gap: Do India's State Drug Buyers Share Bad-Quality Findings?

*A review of public blacklist records from 9 Indian government drug and
medical-supply buyers, 2026-11-07.*

Buyers may have reasons not visible in public records; the firms and buyers
were not contacted for comment.

---

## 1. The question

When one Indian government drug buyer bans a supplier for bad quality, do
other buyers keep it eligible? India has no single national blacklist for
drug and medical-supply suppliers. Each state procurement corporation, and
some national bodies, keep their own separate list, built from their own
inspections and their own test results. This review asks a narrow, factual
question: using only what each buyer has already published, can the same
firm appear as "ineligible" on one buyer's list and "currently supplying"
on another's, for what looks like the same underlying quality problem?

The review centres on Rajasthan Medical Services Corporation Ltd (RMSC),
because its own tender documents contain an eligibility clause that is
explicit about this exact situation: a firm's being found "Not of Standard
Quality" (NSQ) by **any other** state or central government agency makes it
ineligible to bid, not just a finding by Rajasthan itself. That gives this
review a specific, textual test to check each case against, rather than a
general impression.

Every finding below is a possible breach of a tender clause **as worded** -
not a finding of fraud, illegality, or wrongdoing by any firm or buyer.

---

## 2. Data

**9 buyers, 804 ban records**, collected from each buyer's own public
website or published PDF lists between 2026-10-30 and 2026-11-07:

| Buyer | State / scope | Rows | Date range of records | Format |
|---|---|---|---|---|
| APMSIDC | Andhra Pradesh | 43 | 2017-11-29 to 2024-09-10 | PDF (table extraction) |
| AMSCL | Assam | 0 | - | Page's own content is empty server-side (confirmed) |
| PMBI | National (Pradhan Mantri Bhartiya Janaushadhi Pariyojana) | 75 | 2018-09-24 to 2025-04-09 | HTML tables (current + 2 archive pages) |
| GMSCL | Gujarat | 99 | 2021-12-21 to 2026-10-05 | HTML (ASP.NET) |
| MPPHSCL | Madhya Pradesh | 378 | 2016-09-24 to 2026-09-05 | HTML (Firm/Lab/Product lists) |
| RMSC | Rajasthan | 101 | not dated on the live page | HTML |
| KMSCL | Kerala | 5 | 2020-11-29 to 2023-04-27 | PDF |
| KSMSCL | Karnataka | 22 | 2021-02-08 to 2025-10-28 | 3 PDFs, scanned images, hand-transcribed by the project's own researcher directly from the page images (not OCR) |
| BMSICL | Bihar | 81 | not dated (order PDFs are scanned images) | WordPress feed |

**How it was collected**: direct HTTP requests to each buyer's own public
pages/PDFs, respecting robots.txt throughout (one buyer, Gujarat's archive
view, is blocked by a CAPTCHA and was not retrieved; the Andhra Pradesh
`/ftp/` folder is blocked by robots.txt and its files were saved by hand
instead). No login, no paywall, and no access control was bypassed anywhere
in this project.

**What was checked by eye, not just extracted automatically**: the lab/
authority named in 6 RMSC order letters and 4 Madhya Pradesh order letters
(determining whether each is a genuine statutory finding - see Method,
below); the exact wording of RMSC's own eligibility clause and affidavit,
read directly from the tender PDF pages; Karnataka's 3 scanned PDFs,
transcribed in full by hand (not OCR, since a first attempt at OCR could not
reliably keep its dense tables aligned); the Annexure VIII quantity table
for 2 further RMSC tenders that are themselves scanned images, for the 2
firms whose contract lines needed a value estimate; one Madhya Pradesh order
pair confirmed, by a file-hash comparison, to be the same document appearing
under two different reference numbers on that buyer's own site.

---

## 3. Method

**Name matching.** The same firm is often written differently buyer to
buyer ("M/s. Unicure (India) Ltd." vs "Unicure India Ltd"). Every firm name
is run through one shared cleanup step (capitalisation, removing "M/s",
standardising "Pvt"/"Private", "Ltd"/"Limited") before any 2 buyers' records
are compared. Pairs that still look like the same firm after that but are
not textually identical are shown as a separate, human-reviewed candidate
list (`match_candidates.xlsx`) - nothing is ever merged automatically.

**Q1/Q2 quality levels.** Not every "quality" ban is the same kind of
finding. This review splits them in two:
- **Q1** - the ban record names a **statutory** testing authority: a
  Government Analyst, a State Drug Testing Laboratory, CDSCO, a Central
  Drugs Laboratory, or a state Food and Drug (Testing) Laboratory. This is
  a formal "Not of Standard Quality" declaration under Section 25 of the
  Drugs and Cosmetics Act, 1940.
- **Q2** - the quality failure was found at the buyer's own or an
  empanelled lab, or no lab is named at all. This is still a genuine
  quality signal, but not a formal statutory declaration.

Headline findings in this review use **Q1 only**. Q2 cases are reported
separately throughout, always labelled "quality failure in substance, not a
formal NSQ declaration (no government lab named)", and never folded into
the Q1 count.

**Rate-contract comparison.** For each buyer's ban, this review checks
whether the SAME firm also held a live supply ("rate") contract with RMSC
in Rajasthan at the time, by matching the firm's standardised name against
RMSC's own published rate-contract list (2,084 lines, scraped the same way
as the ban lists) and checking date overlap between the other buyer's ban
period and RMSC's own tender bid deadline.

**Rajasthan's clause (ii) reading.** RMSC's drug tenders (clause 5(l)) and
its sutures/surgicals tender (clause 5(k)) carry near-identical wording,
read directly from the tender PDF pages (NIB-10/2024, page 9-10; NIB-10/2026
sutures, page 9):

> *"The concern/company/firm which stands blacklisted/banned/debarred on any
> ground either by Bid Inviting Authority (RMSCL) or Govt. of Rajasthan or
> its departments on the date of bid submission, shall not be eligible to
> participate in the Bid. The concern/company/firm which stands
> blacklisted/banned/debarred on the ground of conviction by court of law or
> the products being found Not of Standard Quality (NoSQ) by any other
> State/Central Government or it's any agencies (central Drugs procurement
> agencies) shall also not be eligible to participate in the Bid. For
> Specific cases regarding other quality issues the purchase committee of
> RMSCL may decide the case on merit basis."*
> - clause 5(l)(II) (drugs) / 5(k)(ii) (sutures)

Read on its own words, this clause disqualifies the **firm**, from the
**whole bid**, once **any** of its products has been found NSQ by another
state or central government agency - it does not require the specific
product banned elsewhere to be the same product being bid on in Rajasthan.
A narrower, product-only reading exists elsewhere in the same tender
(clauses 5(l)(I)/(III) and 5(k)(iii), which bar a bid only "for those
products also" found NSQ elsewhere) - this review reports both readings
side by side for every case, rather than picking one silently.

A court **conviction** is also, on its own, sufficient grounds under this
clause, independent of any lab finding.

---

## 4. Findings

### a) The 3 headline cases

**Zee Laboratories Ltd** - Madhya Pradesh's MPPHSCL banned 3 of Zee's
injectable products as Not of Standard Quality, each confirmed by the
Central Drugs Laboratory, Kolkata (a statutory lab - Q1): Dexamethasone
Sodium Phosphate Injection (order 4102, 11-06-2024), Gentamicin Injection
(order 4790, 04-07-2024), and Ceftriaxone Injection (order 7882,
07-11-2024). Each ban runs 2 years from its order date. Zee separately
holds 73 RMSC rate-contract lines across 3 Rajasthan drug tenders -
NIB-10/2024 (bid deadline 08-10-2024, 13 lines), NIB-14/2024 (deadline
24-02-2025, 34 lines), and NIB-01/2025 (deadline 05-06-2025, 26 lines).
The Madhya Pradesh bans were active on all 3 of those deadlines. Estimated
2-year contract value for these 73 lines: roughly **Rs 25 crore to Rs 46
crore** (see Finance, below, for what the range means and why).

**Meril Endo Surgery Pvt Ltd** - Gujarat's GMSCL debarred Meril's Black
Braided Silk surgical suture (size 1-0) after the Food and Drug Laboratory,
Vadodara (a statutory lab - Q1) reported it Not of Standard Quality
(sterility test), effective 16-07-2025 for 3 years. Meril holds 15 RMSC
rate-contract lines under the Rajasthan sutures tender NIB-10/2026 (bid
deadline 06-05-2026, confirmed by the researcher reading the tender PDF
directly) - active on that deadline. Estimated 2-year contract value for
these 15 lines: roughly **Rs 1 crore to Rs 2.3 crore**.

**Nestor Pharmaceuticals Ltd** - Andhra Pradesh's APMSIDC banned Nestor's
Iron and Folic Acid Syrup, found Not of Standard Quality by the Drugs
Control Laboratory, Vijayawada (a statutory lab - Q1, 4 batches),
effective 06-11-2023 for 2 years. Nestor holds 4 RMSC rate-contract lines
under the Rajasthan drug tender NIB-14/2024 (same deadline as above,
24-02-2025) - active on that deadline. Estimated 2-year contract value for
these 4 lines: roughly **Rs 3 crore to Rs 5.3 crore**.

In every one of these 3 cases, clause 5(l)(II)/5(k)(ii) read as worded
(section 3, above) would have made the firm ineligible to bid on the
relevant Rajasthan tender - a **possible breach of the tender clause as
worded**, not a finding that any rule was actually broken, since this
review does not know what each firm declared in its own bid-time affidavit
(see Limits).

### b) Policy gaps

Beyond the 3 headline cases, this review found 2 further ways the written
rules themselves leave room for a banned firm to keep supplying.

**The Zee CMC eye-drops case.** PMBI (the national Jan Aushadhi body)
debarred Zee Laboratories for Carboxy Methyl Cellulose Eye Drops IP 0.5%,
"Not of Standard Quality", from 18-12-2024. No lab is named on PMBI's own
page, and no order document exists to check further (PMBI's site exposes no
document link for this entry at all) - so this stays Q2, not Q1. Zee
separately holds an RMSC rate-contract line for the exact same product
(Carboxymethylcellulose Eye Drops IP 0.5%), running 31-05-2025 to
31-12-2026 - starting **5 months after** PMBI's debarment began. Because
the PMBI ban started after Rajasthan's own bid deadline (08-10-2024), it
does not make Zee ineligible under clause 5(l)(II) at bid time - the gap
opens only once the Rajasthan contract is already running, mid-contract.

**The affidavit doesn't ask the same question the clause answers on.**
RMSC's own bid-time affidavit (Annexure VII) asks a bidder to swear, on 2
separate points, that it is not blacklisted "on the ground of conviction...
or **the products being found spurious or adulterated**" by any other
government agency - not "Not of Standard Quality", the actual ground
clause 5(l)(II) disqualifies on. "Spurious" and "adulterated" are specific,
narrower legal terms (misbranding or counterfeit) than "Not of Standard
Quality" (a failed lab test, which can include many causes short of
counterfeiting). A firm could, in principle, answer the affidavit's own
question honestly while still being ineligible under the clause's own
separate wording.

**Mid-contract cancellation only covers the narrower ground, too.** RMSC's
own tender document (General Conditions, item (o), page 25) requires a
supplier to self-report any new ban within 15 days of being banned
elsewhere - but only lets RMSC actually **cancel** the running contract if
that other ban's own ground was "spurious or adulterated", the same
narrower wording as the affidavit. A new NSQ finding elsewhere, during the
life of an existing Rajasthan contract, is therefore not itself grounds to
cancel that contract under this document - the Zee CMC case above fits this
gap exactly.

### c) Bans stop at the border

Nothing in the buyer records this review collected states in so many words
that a ban has no effect outside its own state - but 2 pieces of direct
evidence point the same way. Gujarat's own debarment order for Meril
(order no. GMSCL/SD/Debar/23/652/2025-26, read as native PDF text, not
OCR) carries its own distribution list at the bottom of the order - every
single office it is copied to is a Gujarat state office or a GMSCL
warehouse (the Gujarat Principal Secretary for Health, the Gujarat Food &
Drug Control Commissioner, GMSCL's own Ahmedabad/Rajkot/Vadodara/
Bhavnagar/Surat depots, and so on) - no other state, no RMSC, no national
body appears anywhere on that list. Separately, RMSC's own clause
5(l)(II) exists specifically to reach "any other State/Central Government"
ban - a clause that would not need to say this at all if bans already had
effect outside the state that issued them by default. This review could
not find an equivalent explicit distribution list in a readable Madhya
Pradesh order (every MPPHSCL order PDF this review could open is a scanned
image), so this specific finding rests on the Gujarat order and the
structural point about RMSC's own clause, not a second, independent
MPPHSCL document - marked here as a gap in what could be directly checked,
not claimed beyond what the evidence supports.

### d) Concealment cases

5 cases in this review's own data are not about a buyer missing another
buyer's ban, but about a firm's **own declaration** being found false by a
buyer that checked:

| Firm | Caught by | What was concealed |
|---|---|---|
| Unicure India Ltd | GMSCL (Gujarat) | "suppression/non disclosure of debarment" |
| Divine Laboratories Pvt Ltd | GMSCL (Gujarat) | "suppression/non disclosure of debarment" |
| Shiv Paramount Organic And Fertilizers Pvt Ltd | GMSCL (Gujarat) | "Submission of Wrong/false/misleading Self-Attested Declaration of Debarment" |
| Medipol Pharmaceuticals India Pvt Ltd | PMBI (national) | PMBI's own revocation names "Concealing the fact about blacklisting of the firm by GMSD" (a different procurement authority) |
| Radiant Parenterals Ltd | MPPHSCL (Madhya Pradesh) | Did not declare that Tamil Nadu (TNMSC) had blacklisted its Phenytoin Sodium Injection on quality grounds; false affidavit |

These are the buyer's own stated reasons, as written in each order - not
this review's own assessment that a declaration was false.

### e) Court stays

Of the 22 Karnataka (KSMSCL) records in this review - the only buyer whose
full list was hand-transcribed from scanned page images rather than
extracted automatically - **9 are currently stayed by the Karnataka High
Court** (1 firm-level stay, 8 product-level stays), mostly Paschim Banga
Pharmaceuticals (5 separate stay orders across 4 different writ petitions)
and Vivek Pharmachem India Ltd (2 stay orders). *A note on this figure: an
earlier internal summary of this same Karnataka data said "7 of 13" - that
figure came from a narrative count in the source document and does not
match the detailed, row-by-row transcription this review is actually built
on (22 distinct rows, 9 stayed); the "9 of 22" figure here is the one
checked directly against the transcription.*

Across all 9 buyers combined, 40 ban records carry a genuine court stay,
revocation, or (in one case found by this review) an internal
"superseded by a later appeal" correction - concentrated at GMSCL (21 of
the 40) and KSMSCL (9 of the 40). A plain note that a listed ban's period
has simply run out with time is not counted here - that is a different
fact (the ban was real and never overturned) from a stay or revocation.

### f) Network, summary numbers, and the contagion test

Of 369 firms with at least one active ban from some buyer, 51 have an
active ban from exactly 2 different buyers, and 12 from 3 or more -
these 63 firms are shown in the network chart (`output/analysis/
network.png`). Reason categories across all 9 buyers: 371 Quality rows, 106
Non-supply, 36 Other, and 282 still marked "To check" (a reason this review
could not confidently classify from the buyer's own published text alone).

A further, separate test: for firms that also appear in the researcher's
own, independent Project 1 dataset (6,835 CDSCO National Not-of-Standard-
Quality alerts, 2019-2026), this review counted CDSCO alerts in the 12
months before versus after the firm's earliest quality ban in this
project's own data. 94 firms matched between the two datasets (comfortably
over the 20-firm floor below which no conclusion would be drawn). The
aggregate count rose from 202 alerts (before) to 314 (after) - but this is
**not** a clean or uniform result: 37 of the 94 firms rose, 34 fell, and 23
were unchanged, and of this review's own 3 headline cases, only Zee and
Nestor even matched CDSCO's drug-only alert data at all (Meril's ban is for
a surgical suture, outside CDSCO's scope) - and **both fell** after their
ban (Zee 4→3, Nestor 5→0). The aggregate rise is driven by other, unrelated
firms, not this review's own 3 cases. The CDSCO project's own documentation
separately notes a system-wide jump in alert *volume* from increased
state-laboratory reporting from 2025 onward - "do not read the jump as a
change in quality" - a plausible confound this count does not correct for.
**This is a correlational count over a self-selected set of firms (every
one already flagged for a quality problem somewhere), not evidence that any
ban caused any later change.**

### g) Outside evidence

One further, independently confirmed case, found through a general web
search rather than supplied with this project's own source material: a
Comptroller and Auditor General (CAG) of India report, "Public Health
Infrastructure and Management of Health Services", tabled in the Haryana
Vidhan Sabha and reported by The Tribune (14-11-2024), found that Haryana
Medical Services Corporation Ltd purchased drugs worth an estimated Rs 1.52
crore from Nestor Pharmaceuticals between May and December 2019 - **after**
Gujarat's GMSCL had already blacklisted Nestor for 3 years in February 2017
for the same drug (folic acid-ferrous sulphate tablets), and after Haryana
itself separately blacklisted Nestor in September 2018 for concealing that
earlier Gujarat action. This is a different buyer pair (Gujarat/Haryana,
not Gujarat/Rajasthan) from this review's own Nestor finding above, and was
independently confirmed by the CAG's own audit, not derived from this
project's data.

---

## 5. Finance

Every figure in this section is an **estimated contract value (2-year),
not amount paid** - built from Rajasthan's own tender-time quantity
estimate (Annexure VIII) multiplied by the rate the firm was contracted at,
never from an actual purchase or delivery record. 2 estimates are shown for
every contract line, because Annexure VIII's quantity figure is the
tender's total across every supplier of that drug, not a per-supplier
figure, when more than one firm holds a rate contract for the same drug
code:

- **Upper estimate** - the full tender quantity assigned to this one
  supplier (an upper bound).
- **Equal-share estimate** - the upper estimate divided evenly across every
  supplier holding a rate contract for that same drug code under that same
  tender.

| Firm | Upper estimate (2-year) | Equal-share estimate (2-year) | Lines with a computed value |
|---|---|---|---|
| Zee Laboratories Ltd | Rs 46.37 crore | Rs 25.28 crore | 56 of 73 (14 lines could not be matched to a tender's own Annexure VIII entry, or had an unreadable pack-size instruction; 3 further lines, in the scanned NIB-01/2025 tender, were hidden behind a digital-signature stamp and were left blank, not guessed) |
| Meril Endo Surgery Pvt Ltd | Rs 2.28 crore | Rs 0.99 crore | 15 of 15 |
| Nestor Pharmaceuticals Ltd | Rs 5.27 crore | Rs 3.08 crore | 3 of 4 |

2 of Rajasthan's 4 relevant tenders (NIB-10/2024, NIB-14/2024) are native,
machine-readable PDF text; their own quantity figures were read directly by
this project's own extraction code. The other 2 (NIB-01/2025, NIB-10/2026)
are scanned images - their quantity tables were read by the project's
researcher directly from the page images, not by OCR (an OCR attempt on
these same scanned tables could not reliably keep columns aligned, for the
same reason already documented for Karnataka's scanned lists above). Every
figure above, and every line behind it, is in `output/finance/
headline_case_values.xlsx`.

---

## 6. International comparison

The United States runs its exclusions the opposite way: SAM.gov (System
for Award Management), operated by the General Services Administration, is
a single, government-wide database - when **any** federal agency debars,
suspends, or otherwise excludes a contractor, that action flows directly
into the one shared list, and every other federal agency checks the same
list before awarding its own contracts. The FDA separately keeps its own
Debarment List (under the Federal Food, Drug, and Cosmetic Act, added by
the Generic Drug Enforcement Act of 1992) specifically for drug-application
integrity - a second, narrower list alongside SAM.gov's own broader
procurement-exclusion one, but still a single national list for its own
purpose, not a separate list per US state. India's state procurement
corporations, by contrast, each keep their own, separately published
blacklist, built from their own inspections - there is no equivalent
shared registry a Rajasthan or Haryana buyer could check before awarding a
contract; finding out whether a firm bidding in one state has already been
banned elsewhere, as this review did, currently requires checking each
buyer's own list by hand, one at a time.

*Sources: [US Federal Exclusions (SAM.gov)](https://apify.com/faceless-api/us-federal-exclusions-sam-gov);
[What Is the FDA Debarment List?](https://streamlineverify.com/glossary/sanctions/fda-debarment-list/);
[OIG, SAM.gov, & State Exclusion Lists Explained](https://providertrust.com/blog/oig-sam-state-medicaid-exclusion-lists-differences);
[US SAM Procurement Exclusions](https://www.opensanctions.org/datasets/us_sam_exclusions/).*

---

## 7. Limits

- **OCR.** Several source documents (RMSC's Hindi-language order letters,
  2 RMSC tenders' quantity tables, Karnataka's 3 blacklist PDFs, most
  Madhya Pradesh order letters) are scanned images, not machine-readable
  text. Where OCR was used, every resulting value is marked "OCR - check"
  and was never treated as confirmed on its own; where OCR could not
  reliably read a dense table at all (the 2 scanned RMSC tenders'
  quantities, Karnataka's full detail), the project's own researcher read
  the figures directly from the page images instead.
- **9 of 30+ buyers.** India has many more state and national drug/medical-
  supply procurement bodies than the 9 reviewed here. 3 further buyers
  (Telangana's TGMSIDC, Odisha's OSMCL, Tamil Nadu's TNMSC) were attempted
  and found blocked (no public blacklist page found, DNS failure, and a
  consistently timing-out listing page, respectively) - not included.
  Assam's own page returned no content at all (confirmed, not a collection
  bug).
- **Missing archives.** Gujarat's own "Archive Blacklisted" view is blocked
  by a CAPTCHA to automated access and was not collected - only Gujarat's
  current list is included.
- **Product vs firm reading.** RMSC's tender text itself is genuinely
  ambiguous between a whole-firm and a narrower, product-only
  disqualification reading (section 3). This review reports both; a court
  or RMSC's own purchase committee could reasonably read it either way.
- **Upper estimates, not amounts paid.** Every finance figure in section 5
  is a tender-time quantity estimate, not a record of money actually
  transferred - see that section for the full caveat.
- **No purchase/delivery data.** This review has no record of which
  Annexure VIII quantities were actually ordered, delivered, or paid for -
  only what RMSC estimated at tender time and what rate each firm was
  contracted at.
- **Firms and buyers not contacted.** No firm or buyer named in this review
  was contacted for comment. Buyers may have reasons not visible in public
  records - a disclosed affidavit this review has not seen, an internal
  waiver, or a decision by the purchase committee under the clause's own
  "merit basis" language (section 3) - that would change how a specific
  case should be read.

---

## 8. Recommendations

1. **A shared national blacklist.** A single, central, publicly searchable
   registry of pharmaceutical and medical-supply debarments - covering
   every state and national procurement body, built on the SAM.gov model
   described in section 6 - would let any buyer check a bidder's full
   record in one place, rather than requiring the buyer-by-buyer search
   this review had to carry out by hand.
2. **One common clause wording**, used by every state buyer, that clearly
   and explicitly reaches a Not-of-Standard-Quality finding by *any* other
   procurement agency - removing the product-only/whole-firm ambiguity this
   review found in Rajasthan's own tender text (section 3), and removing
   the need for each state to draft its own version.
3. **Align the affidavit and the cancellation clause to the eligibility
   ground.** Rajasthan's own bid-time affidavit, and its mid-contract
   cancellation clause, currently ask about a narrower ground ("spurious or
   adulterated") than the eligibility clause actually disqualifies on
   ("Not of Standard Quality") - see section 4(b). Matching all 3 to the
   same NSQ wording would close the specific gap the Zee CMC eye-drops case
   shows in practice.

---

## 9. Sources

Every PDF, webpage, and spreadsheet cited in this report, with the page or
sheet it was read from:

- RMSC drug tender NIB-10/2024: `data/raw/RMSC_Tenders/NIB10_2024_Tender_NIb10_DM.pdf`, clause 5(l), pages 9-10; Annexure VIII, pages 60-71 - https://rmsc.health.rajasthan.gov.in/
- RMSC drug tender NIB-14/2024/69: `data/raw/RMSC_Tenders/NIB14_2024_69_Tender14.pdf`, Annexure VIII, pages 59-95
- RMSC drug tender NIB-01/2025/871: `data/raw/RMSC_Tenders/NIB01_2025_NIB01.pdf` (scanned), Annexure VII affidavit pages 55-56, Annexure VIII pages 59-82
- RMSC sutures tender NIB-10/2026: `data/raw/RMSC_Tenders/NIB10_2026_Sutures_Finalbid10UBN.pdf` (scanned), clause 5(k), page 9; Annexure VIII (List of Item(s) with Specifications), pages 54-58
- RMSC General Conditions, item (o): same tender documents, page 25
- RMSC's own rate-contract list: `output/frozen_2026-11-07/rate_contracts/RMSC.xlsx` - https://rmsc.health.rajasthan.gov.in/content/raj/medical/rajasthan-medical-services-corporation-ltd-/en/Approved-Rate-Lists/DrugsRC.html
- MPPHSCL ban orders for Zee (4102, 4790, 7882, 8078) and Meril (7876/166):
  `output/frozen_2026-11-07/buyers/MPPHSCL.xlsx` - https://mpphscl.mp.gov.in/
- GMSCL debarment order for Meril (GMSCL/SD/Debar/23/652/2025-26):
  `data/raw/GMSCL/candidate_gap_pdfs/GMSCL-DEBAR-18621-652-2025-26.pdf`
- APMSIDC blacklisted items list (Nestor): `data/text/AP_items_blacklisted_2024-09-13.txt` - https://apmsidc.ap.nic.in/
- PMBI debarment list (Zee CMC eye drops): `output/frozen_2026-11-07/buyers/PMBI.xlsx` - https://www.pmbi.co.in/debar.aspx
- KSMSCL (Karnataka) 3 scanned PDFs, hand-transcribed: `output/frozen_2026-11-07/buyers/KSMSCL.xlsx`
- Concealment cases: `output/frozen_2026-11-07/concealment_cases.xlsx`
- CAG Haryana report ("Public Health Infrastructure and Management of
  Health Services") via The Tribune, 14-11-2024 -
  https://www.tribuneindia.com/news/haryana/haryana-bought-drugs-from-blacklisted-firm-cag/
- This project's own data-quality log: `output/source_issues.csv`,
  `PROGRESS_LOG.md`
- Finance figures: `output/finance/headline_case_values.xlsx`
- Network/summary/contagion figures: `output/analysis/network.png`,
  `output/analysis/summary_numbers.xlsx`, `output/analysis/contagion_test.xlsx`
- Project 1 (CDSCO NSQ) dataset: `E:\cdsco_mqri\final_package\CDSCO_NSQ_Project_Workbook.xlsx`
  (the researcher's own separate, prior project)
- US comparison: see section 6 above for individual links.
