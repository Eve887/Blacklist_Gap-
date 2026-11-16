"""
common.py - small helper tools shared by every buyer script.

You never run this file directly. The other scripts "import" it,
which means they borrow the functions written here.
"""

import re                      # "re" = regular expressions: a way to search text by pattern
from pathlib import Path       # Path = a safe way to build folder and file paths on Windows
from collections import defaultdict

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ------------------------------------------------------------------
# FOLDERS
# __file__ is the path of this file. .parent is the folder it sits in
# (scripts), and .parent again is the project folder (blacklist_gap).
# Building paths this way means the scripts work no matter which
# folder Command Prompt is "standing in" when you run them.
# ------------------------------------------------------------------
PROJECT = Path(__file__).resolve().parent.parent
RAW = PROJECT / "data" / "raw"            # downloaded PDFs (never edited)
TEXT = PROJECT / "data" / "text"          # plain text copied out of each PDF
OCR_TEXT = TEXT / "ocr"                   # OCR output only (PROJECT_RULES.md rule F) - never mixed with
                                           # normal extracted text, so it's always obvious which
                                           # folder a given .txt file came from
OUT = PROJECT / "output"                  # Excel files the scripts produce
BUYERS = OUT / "buyers"                   # one Excel file per buyer; 10_merge_master.py combines them

# Tesseract's own install does not reliably land on PATH (confirmed on
# this machine - see PROGRESS_LOG.md's OCR-reinstall entries), and PATH
# changes don't reliably propagate to every way a script might get run
# (a fresh terminal, a scheduled task, this project's own tooling) -
# so the full path is set explicitly here instead, the same kind of
# narrow, explicit fix already used for MPPHSCL's TLS issue rather than
# relying on environment state this project doesn't control.
TESSERACT_PATH = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")

for folder in (RAW, TEXT, OUT, BUYERS):
    folder.mkdir(parents=True, exist_ok=True)   # create the folder if missing; do nothing if it exists


# ------------------------------------------------------------------
# THE 19 MASTER-LIST COLUMNS
# Every buyer script (03_build_master.py, 03_assam_master.py, 04_pmbi.py, ...)
# writes its own output\buyers\<BUYER>.xlsx using exactly these columns, so
# 10_merge_master.py can combine them all into one master_list.xlsx safely.
# ------------------------------------------------------------------
COLUMNS = ["Entry ID", "Firm name (as written)", "Firm name (standard)", "Firm address", "Buyer",
           "Buyer state", "Level (firm / product)", "Product(s)", "Blacklisted from", "Blacklisted to",
           "Period", "Reason (as written)", "Reason category (Quality / Non-supply / Other / To check)",
           "Lab that declared NSQ", "Court stay or revoked?", "Source link", "Page / row in source",
           "Status", "Date entered"]

# Placeholder every buyer script writes into the "Reason category" column -
# moved out of the buyer scripts 2026-11-01 (PROGRESS_LOG.md Entry 19, per
# checklist item 13): actual classification (reason_category()/
# quality_level(), including the MPPHSCL/RMSC hand-verified overrides) now
# runs ONCE, centrally, in 10_merge_master.py on every merge, so a
# QUALITY_WORDS/SUPPLY_WORDS change takes effect across every buyer the next
# time the merge runs - no buyer script needs a separate re-run, which is
# exactly how the Meril/GMSCL bug happened (Entries 16-17). Buyer scripts now
# only write the raw "Reason (as written)" text and the lab name.
CATEGORY_PENDING = "Pending (computed centrally by 10_merge_master.py)"


# ------------------------------------------------------------------
# DUPLICATE ENTRY IDs
# Entry IDs are built from the source (a serial number, a year + serial, a
# date + code, ...) instead of a running count, so they stay the same from
# one run to the next. If two rows still end up with the same ID (a genuine
# clash in the source numbering), keep the first one as-is and add "-b",
# "-c", ... to the rest - and print a warning so it does not pass unnoticed.
# ------------------------------------------------------------------
def dedup_entry_ids(rows):
    groups = defaultdict(list)
    for r in rows:
        groups[r[0]].append(r)
    warnings = []
    for base_id, group in groups.items():
        if len(group) < 2:
            continue
        for i, r in enumerate(group):
            if i > 0:
                r[0] = f"{base_id}-{chr(ord('b') + i - 1)}"
        warnings.append(f"{base_id} -> " + ", ".join(r[0] for r in group))
    return warnings


# ------------------------------------------------------------------
# MATCHES - same standard name appearing more than once (same buyer or
# different buyers). Used by 10_merge_master.py on the FULL combined list,
# so it finds cross-buyer matches too. Rows whose Level is "Lab" (any
# wording starting with "Lab") are excluded - a lab is a testing facility,
# not a manufacturer, so matching a lab's name against a firm's name would
# be a false "cross-buyer case", not a real one.
# ------------------------------------------------------------------
def find_matches(rows):
    groups = defaultdict(list)
    for r in rows:
        if r[6].strip().lower().startswith("lab"):          # r[6] = Level (firm / product)
            continue
        groups[r[2]].append(r)                    # r[2] = Firm name (standard)
    out = []
    for name, items in sorted(groups.items()):
        if len(items) < 2:
            continue
        buyers = sorted({r[4] for r in items})
        out.append([name, len(items), ", ".join(r[0] for r in items), ", ".join(buyers),
                    "DIFFERENT buyers - possible cross-state/cross-buyer case" if len(buyers) > 1
                    else "Same buyer, repeat action"])
    out.sort(key=lambda x: (not x[4].startswith("DIFFERENT"), x[0]))   # cross-buyer cases first
    return out


# ------------------------------------------------------------------
# WRITE ONE EXCEL SHEET
# Plain typed values, Arial font, no formulas, no merged cells - used by
# every buyer script and by 10_merge_master.py.
# ------------------------------------------------------------------
def write_sheet(ws, header, rows, widths):
    head_font = Font(name="Arial", bold=True, color="FFFFFF")
    fill = PatternFill("solid", start_color="1F4E78")
    side = Side(style="thin", color="BFBFBF")
    border = Border(left=side, right=side, top=side, bottom=side)
    ws.append(header)
    for r in rows:
        ws.append(r)
    for row in ws.iter_rows():
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = border
            c.font = head_font if c.row == 1 else Font(name="Arial", size=10)
            if c.row == 1:
                c.fill = fill
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    ws.freeze_panes = "C2"


# ------------------------------------------------------------------
# NAME STANDARDISING  (the "Firm name (standard)" column)
# Example:  "M/s. Unicure (India) Pvt. Limited"  ->  "UNICURE INDIA PVT LTD"
# ------------------------------------------------------------------
def standard_name(name):
    if not name:
        return ""
    s = name.upper()                                   # 1. all capital letters
    s = re.sub(r"\bM\s*/\s*S\b\.?", " ", s)            # 2. remove "M/s", "M/S.", "M / s"
    s = s.replace("&", " AND ")                        # 3. "&" and "and" become the same
    s = re.sub(r"[^\w\s]", " ", s)                     # 4. remove every punctuation mark . , ( ) ' -
    s = re.sub(r"\bPRIVATE\b", "PVT", s)               # 5. PRIVATE -> PVT
    s = re.sub(r"\bLIMITED\b", "LTD", s)               # 6. LIMITED -> LTD
    s = re.sub(r"\bLABS\b", "LABORATORIES", s)         # 7. LABS -> LABORATORIES
    s = re.sub(r"\s+", " ", s).strip()                 # 8. many spaces -> one space
    return s


# ------------------------------------------------------------------
# SPLIT "name + address" cells
# The Andhra Pradesh PDFs put the firm name and address in ONE cell:
#   "M/s Bhat Bio Tech India\nPvt Ltd\n11 'A', 4th Cross, ..."
#
# Rule 1 (tried FIRST) - "legal ending": look at the first 4 lines, but
# stop early at the first line that is clearly an address (ADDRESS_START -
# a number, '#', 'Plot', 'D.No', 'Sy', 'Survey', 'Khasra', 'Village' or
# 'Vill'). If the words "Ltd", "Limited" or "LLP" show up anywhere in that
# stretch of text, the name is everything up to and INCLUDING the FIRST one
# of those words - even if it falls in the middle of a line (PDFs often
# wrap "...Pvt Ltd" and the start of the address onto the same line) -
# and everything after it is the address.
# We deliberately do NOT look for ordinary business words like
# "Pharmaceuticals" or "Biotech" here, because an address can contain
# those too ("Pharma City", "Biotech Park") and would get wrongly pulled
# into the name.
#
# Rule 2 (fallback, used only when Rule 1 finds no "Ltd"/"Limited"/"LLP"):
# the older line-by-line rules - stop at an ADDRESS_START line, but keep
# the name growing while either (a) the name built so far ends with "And",
# "&" or "Of" (it is clearly not finished), or (b) the next line starts
# with a company word such as "Pharmaceuticals" or "Ltd".
# ------------------------------------------------------------------
ADDRESS_START = re.compile(r"^\s*(\d|#|plot|d\.?\s*no|sy\b|survey|khasra|village|vill\b|s\.?\s*no)", re.I)

LEGAL_SUFFIX = re.compile(r"\b(ltd|limited|llp)\b\.?", re.I)

NAME_ENDS_WITH_JOINER = re.compile(r"(?:\b(?:and|of)|&)\s*$", re.I)

COMPANY_WORD_START = re.compile(
    r"^\s*(pharmaceuticals?|pharma|ltd\.?|limited|pvt\.?|private|laboratories|labs|"
    r"healthcare|remedies|drugs|life\s*sciences|biotech|industries|corporation|llp)\b",
    re.I,
)

def split_name_address(cell):
    if not cell:
        return "", ""
    lines = [l.strip() for l in str(cell).split("\n") if l.strip()]
    if not lines:
        return "", ""

    # Rule 1: build the candidate name text from the first 4 lines, but
    # stop early if a line is clearly the start of the address.
    prefix_lines = [lines[0]]
    for line in lines[1:4]:
        if ADDRESS_START.match(line):
            break
        prefix_lines.append(line)
    candidate_text = " ".join(prefix_lines)
    legal = LEGAL_SUFFIX.search(candidate_text)
    if legal:
        name = candidate_text[:legal.end()]
        after_match = candidate_text[legal.end():]            # leftover text on the same line, if any
        remaining_lines = lines[len(prefix_lines):]            # lines after the ones we looked at
        address_parts = ([after_match] if after_match.strip() else []) + remaining_lines
        return _tidy(name), _tidy(" ".join(address_parts))

    # Rule 2 (fallback): no "Ltd"/"Limited"/"LLP" found - grow the name
    # line by line using the earlier rules.
    name_end = 1
    while name_end < len(lines):
        next_line = lines[name_end]
        if ADDRESS_START.match(next_line):
            break                                     # a clear address marker - stop here
        name_so_far = " ".join(lines[:name_end])
        if NAME_ENDS_WITH_JOINER.search(name_so_far) or COMPANY_WORD_START.match(next_line):
            name_end += 1                              # name is not finished - pull in this line
        else:
            break                                      # neither rule says "keep going" - stop here

    name = _tidy(" ".join(lines[:name_end]))
    address = _tidy(" ".join(lines[name_end:]))
    return name, address


def _tidy(text):
    """Remove doubled spaces, spaces before commas, and a comma or full stop left at the end."""
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+,", ",", text)
    return text.strip(" ,;")


# ------------------------------------------------------------------
# STRIP A LEAKED ADDRESS OUT OF A "FIRM NAME" FIELD
# Some buyer sites (found on GMSCL and MPPHSCL) put the address directly in
# their own "Company Name" / "Name of Firm" field - there is no separate
# address column to split from at all, so split_name_address() (which
# works on PDF cells with real line breaks) does not apply here. This
# looks instead for STRUCTURAL signals that an address has been appended:
#   - a comma followed by a number (a street/plot number)
#   - a comma followed by a recognized city/state name
#   - an explicit address marker word anywhere (Plot No, Add:, Village,
#     Nagar, Road, Dist/District, Estate, Society, Sector, Floor, Highway,
#     Taluka, Survey No, Industrial Area)
#   - a 6-digit PIN code
#   - a recognized city/state name appearing AFTER a legal-entity suffix
#     (Ltd/Limited/Pvt/Private/LLP) has already appeared earlier in the
#     text - catches a trailing place name with no comma at all (e example:
#     "...Pvt Ltd. New Delhi")
# Deliberately narrow: a place name that is NOT after a comma or a legal
# suffix is left alone, so a real company name that happens to include a
# city/state (e.g. "Kerala State Drugs And Pharmaceuticals Ltd", "Baroda
# Surgical (India) Private Limited", "Gujarat Pharmalab Pvt Ltd") is not
# touched - confirmed against this project's own real data, not guessed.
# ------------------------------------------------------------------
ADDRESS_MARKER_RE = re.compile(
    r"\bplot\s*no\.?|\badd\s*:|\baddress\s*:|\bvillage\b|\bnagar\b|\broad\b|\bdist(?:rict)?\b|"
    r"\bestate\b|\bsociety\b|\bsector\b|\bfloor\b|\bhighway\b|\btaluka\b|\bsurvey\s*no\.?|"
    r"\bindustrial\s*area\b", re.I)

PLACE_NAMES = {"mumbai", "navi mumbai", "delhi", "new delhi", "indore", "ahmedabad", "hyderabad",
               "chennai", "bangalore", "bengaluru", "pune", "kolkata", "jaipur", "lucknow", "nagpur",
               "surat", "vadodara", "rajkot", "nashik", "noida", "gurgaon", "gurugram", "faridabad",
               "ghaziabad", "dehradun", "chandigarh", "bhopal", "jabalpur", "ujjain", "ratlam",
               "rajasthan", "maharashtra", "gujarat", "uttarakhand", "madhapur", "hyedrabad", "raigad"}
PLACE_NAME_RE = re.compile("|".join(re.escape(p) for p in PLACE_NAMES), re.I)

PIN_CODE_RE = re.compile(r"\b\d{6}\b")

LEGAL_ENTITY_SUFFIX_RE = re.compile(r"\b(ltd|limited|pvt|private|llp)\b", re.I)


def strip_leaked_address(text):
    """Return `text` with a leaked address trimmed off, or unchanged if no
    address signal is found."""
    if not text:
        return text
    candidates = []
    for m in ADDRESS_MARKER_RE.finditer(text):
        if m.start() > 0:
            candidates.append(m.start())
    for m in re.finditer(r",\s*", text):
        after_comma = text[m.end():]
        if re.match(r"^\d", after_comma) or PLACE_NAME_RE.search(after_comma):
            candidates.append(m.start())
    pin_match = PIN_CODE_RE.search(text)
    if pin_match:
        candidates.append(pin_match.start())
    suffix_matches = list(LEGAL_ENTITY_SUFFIX_RE.finditer(text))
    if suffix_matches:
        place_after_suffix = PLACE_NAME_RE.search(text, suffix_matches[-1].end())
        if place_after_suffix:
            candidates.append(place_after_suffix.start())
    if not candidates:
        return text
    return text[:min(candidates)].rstrip(" ,.")


# ------------------------------------------------------------------
# DATES   "29.11.2017" or "29/11/2017" -> "29-11-2017"
# Kept as TEXT on purpose, so Excel never changes it into a US-style date.
# ------------------------------------------------------------------
DATE_RE = re.compile(r"(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})")

def clean_date(text):
    if not text:
        return ""
    m = DATE_RE.search(str(text))
    if not m:
        return str(text).strip()
    d, mth, y = m.groups()
    if len(y) == 2:
        y = "20" + y
    return f"{int(d):02d}-{int(mth):02d}-{y}"


def looks_like_date(text):
    """True if the text has something date-shaped in it (dd.mm.yyyy, dd/mm/yyyy, dd-mm-yyyy).
    Used as a safety check: a row is only accepted if its From/Up to cells look like real dates."""
    return bool(text) and DATE_RE.search(str(text)) is not None


# ------------------------------------------------------------------
# REASON CATEGORY
# Looks for key words in the reason text and gives one label.
# It is only a FIRST GUESS - you must check every row by reading it.
# ------------------------------------------------------------------
QUALITY_WORDS = ["not of standard quality", "nsq", "nosq", "substandard", "sub-standard", "sub standard",
                 "spurious", "adulterated", "misbranded", "failed in quality", "quality test",
                 "failed the test", "contamina",
                 # Added 2026-10-31 per user's review (specific lab-test/QC terms that show up
                 # in reasons without ever saying "quality" or "NSQ" outright):
                 "dissolution", "disintegration", "sterility", "endotoxin", "assay", "content of",
                 "identification", "related substances", "particulate", "uniformity",
                 "nabl failure", "quality failure", "quality ground"]
SUPPLY_WORDS = ["non supply", "non-supply", "default in supply", "non-completion", "non completion",
                "not supplied", "delay in supply", "failure to supply", "short supply", "non-performance",
                "non performance"]
# "ph" (pH test failure) is a real QC term, but as a plain substring it would match inside
# "pharma"/"pharmaceutical"/"Phenytoin"/"Phenobarbitone" and other ordinary words that have
# nothing to do with pH - checked with a WORD BOUNDARY instead of adding it to QUALITY_WORDS,
# where every other entry is safe to substring-match (none of them are plausible fragments of
# an unrelated word).
PH_WORD_RE = re.compile(r"\bph\b", re.I)

def reason_category(text):
    t = re.sub(r"\s+", " ", (text or "")).lower()     # join words split across lines first
    q = any(w in t for w in QUALITY_WORDS) or PH_WORD_RE.search(t)
    s = any(w in t for w in SUPPLY_WORDS)
    if q and s:
        return "Both - check"
    if q:
        return "Quality"
    if s:
        return "Non-supply"
    return "To check"


# Q1/Q2 split, added 2026-10-31 per the user's instruction. Legal basis
# (recorded here AND in PROGRESS_LOG.md section 3.5, per the user's
# request): under the Drugs and Cosmetics Act 1940, a drug is formally
# declared "Not of Standard Quality" via a Government Analyst's report
# (Section 25) - a STATUTORY finding, distinct from a buyer's own or an
# empanelled lab's test result, even though the latter is often ALSO a
# genuine quality failure in substance (a failed pharmacopoeial test is
# NSQ in substance, just not yet formally declared by the statutory
# process). Q1 = that statutory declaration is named; Q2 = everything
# else that is still a quality signal (buyer's own/empanelled lab, or no
# lab named at all) - never a non-quality row (callers should only call
# this where reason_category() already returned "Quality").
Q1_LAB_WORDS = ["government analyst", "govt. analyst", "govt analyst", "state drug testing",
               "state drug test", "rajkiya aushadhi", "rajkiya", "central drugs laboratory",
               "central drug laboratory", "drugs control", "drug testing laboratory",
               "government lab", "govt. lab", "govt lab", "state lab",
               # Found missing 2026-10-31 (Meril/GMSCL case) - Gujarat's own statutory drug
               # testing body uses a different name entirely ("Food and Drug Laboratory",
               # under its Food and Drug Control Administration) - a real gap, not assumed
               # to be the only one; other states may use still other names not yet seen.
               "food and drug laboratory", "food and drugs laboratory", "fdca"]
# Short acronyms get their own word-boundary check - as plain substrings
# they risk matching inside unrelated text (found no concrete false-
# positive example, but these are short enough to be cautious anyway).
Q1_ACRONYM_RE = re.compile(r"\b(cdsco|cdtl|dcl|cdl)\b", re.I)


def quality_level(reason_as_written, lab_that_declared_nsq, reason_category_value):
    """Returns "Q1" (a named statutory/government lab), "Q2" (buyer's
    own/empanelled lab or no lab named - still a genuine quality
    signal), or "N/A" if `reason_category_value` isn't "Quality" (never
    guessed for a non-quality row)."""
    if reason_category_value != "Quality":
        return "N/A"
    combined = re.sub(r"\s+", " ", f"{reason_as_written or ''} {lab_that_declared_nsq or ''}").lower()
    if any(w in combined for w in Q1_LAB_WORDS) or Q1_ACRONYM_RE.search(combined):
        return "Q1"
    return "Q2"


def one_line(value):
    """Turn a cell with line breaks into one tidy line of text."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


# ------------------------------------------------------------------
# OCR (PROJECT_RULES.md rule F) - used only for scanned/image PDFs where
# pdfplumber's normal extract_text() returns nothing. Confirmed working
# on this machine 2026-10-31 (Tesseract 5.5.3, official UB-Mannheim/
# tesseract-ocr Windows build - see PROGRESS_LOG.md for the reinstall).
# ------------------------------------------------------------------

def ocr_page_text(pdf_page, resolution=300, lang="eng"):
    """Renders one pdfplumber page to an image and OCRs it. Returns the
    raw OCR text - NOT checked, NOT trusted for any finding on its own
    (rule F) - the caller must still save it via save_ocr_text() and
    mark every value taken from it "OCR - check". `lang` matches
    Tesseract's own -l syntax (e.g. "hin+eng" for a document mixing
    Hindi and English, as RMSC's order letters do - see
    PROGRESS_LOG.md). Needs that language's .traineddata file in
    tessdata - confirmed present via `tesseract --list-langs` before
    adding Hindi support (2026-10-31)."""
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = str(TESSERACT_PATH)
    image = pdf_page.to_image(resolution=resolution).original
    return pytesseract.image_to_string(image, lang=lang)


def save_ocr_text(source_pdf_name, page_number, text):
    """Saves OCR output to data/text/ocr/<file>_p<page>.txt, per
    PROJECT_RULES.md rule F - OCR text always lives next to the source PDF it
    came from, never only inside a spreadsheet cell. Returns the path."""
    OCR_TEXT.mkdir(parents=True, exist_ok=True)
    stem = Path(source_pdf_name).stem
    out_path = OCR_TEXT / f"{stem}_p{page_number}.txt"
    out_path.write_text(text, encoding="utf-8")
    return out_path
