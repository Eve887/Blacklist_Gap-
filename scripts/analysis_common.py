"""
analysis_common.py - shared helpers for the Part 1 "data science" scripts
(17_network_chart.py, build_summary_numbers.py, build_contagion_test.py,
build_analysis_charts.py), per the user's FINAL STAGE instruction to use
ONLY output/frozen_2026-11-07/ (plus output/finance/ and output/review/)
for every fact in this stage - never the live output/ folder.

Never run directly.
"""

import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from common import PROJECT

FROZEN = PROJECT / "output" / "frozen_2026-11-07"
ANALYSIS_DIR = PROJECT / "output" / "analysis"
ANALYSIS_DIR.mkdir(parents=True, exist_ok=True)

# The 3 headline cases, standardized names (match_decisions.xlsx/case
# files' own spelling) - highlighted throughout Part 1.
HEADLINE_FIRMS = {"ZEE LABORATORIES LTD", "MERIL ENDO SURGERY PVT LTD",
                  "NESTOR PHARMACEUTICALS LTD"}

# dataviz skill palette (categorical slots, fixed order - never cycled;
# see references/palette.md) - used by every PNG chart in Part 1/2.
PALETTE = {
    "blue": "#2a78d6", "orange": "#eb6834", "aqua": "#1baf7a",
    "yellow": "#eda100", "magenta": "#e87ba4", "green": "#008300",
    "violet": "#4a3aa7", "red": "#e34948",
}
INK_PRIMARY = "#0b0b0b"
INK_SECONDARY = "#52514e"
INK_MUTED = "#898781"
GRIDLINE = "#e1e0d9"
BASELINE = "#c3c2b7"
SURFACE = "#fcfcfb"

# A row's "Court stay or revoked?" field carries 3 genuinely different
# things, as written (see PROGRESS_LOG.md, 10_merge_master.py/
# 07_mpphscl.py): (1) "Not stated" - nothing unusual; (2) a GENUINE
# judicial stay, revocation, or (new 2026-11-07) this project's own
# "superseded by appeal" finding - the ban is NOT a live, un-overturned
# action; (3) a plain "Blacklisting Period Completed" note - the ban was
# real and was never overturned, it has simply expired with time, which
# is a different fact entirely (every old ban eventually says this).
# "Active ban" throughout Part 1 means (1) or (3) - a real, never-
# overturned ban - excluding only (2). This is the SAME test
# build_rc_vs_bans.py's classify() already uses for its Group "X"
# (stayed/revoked) exclusion, just generalized to cover the new
# "SUPERSEDED" wording too.
EXCLUDE_COURT_RE = re.compile(
    r"\bstay(ed)?\b|\brevoked?\b|\brevocation\b|\bsuperseded\b|\bset[\s-]?aside\b|"
    r"\bsetting aside\b|\babeyance\b|\bsuspension\b|\binterim\b", re.I)

EXCLUDE_LEVELS = {"Lab"}                  # labs are not manufacturers/firms


def is_active_ban(court_text, level):
    if level in EXCLUDE_LEVELS:
        return False
    court_text = (court_text or "").strip()
    if court_text and court_text.lower() != "not stated" and EXCLUDE_COURT_RE.search(court_text):
        return False
    return True


DATE_FORMATS = ("%d-%m-%Y", "%d-%b-%Y", "%d-%B-%Y")


def parse_date(text):
    """Returns a datetime.date, or None if unparseable - never guessed.
    master_list.xlsx mixes 2 date spellings across buyers (03_build_master.py
    "29-11-2017" vs 07_mpphscl.py "06-Nov-2024") - both tried."""
    if not text or str(text).strip().lower() in ("not stated", ""):
        return None
    text = str(text).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def load_master_rows():
    """Returns (header, rows) from output/frozen_2026-11-07/master_list.xlsx
    "Master list" sheet - the ONLY source of buyer/firm/ban facts for
    every Part 1 script, per the user's instruction."""
    wb = load_workbook(FROZEN / "master_list.xlsx", read_only=True, data_only=True)
    ws = wb["Master list"]
    rows = list(ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    idx = {h: i for i, h in enumerate(header)}
    data = [r for r in data if any(c is not None and str(c).strip() for c in r)]
    return idx, data


def load_concealment_cases():
    wb = load_workbook(FROZEN / "concealment_cases.xlsx", read_only=True, data_only=True)
    ws = wb["Concealment cases"]
    rows = list(ws.iter_rows(values_only=True))
    header, data = rows[0], rows[1:]
    idx = {h: i for i, h in enumerate(header)}
    return idx, data
