"""
build_analysis_charts.py - Part 1 (data science), FINAL STAGE. READ-ONLY
REPORT: 4 PNG charts, using ONLY output/frozen_2026-11-07/ (plus the
already-confirmed RMSC tender deadlines, same constants
build_rc_vs_bans.py uses - not re-derived here).

Color/layout follows the dataviz skill's reference palette
(analysis_common.PALETTE) - fixed categorical hue order, one axis per
chart, a legend for every multi-series chart, a labelled source line
under each PNG (never left for the reader to guess where a number came
from).

Writes to output/analysis/:
  bans_by_buyer_reason.png   - bans per buyer, by reason category
  q1_vs_q2.png                - Q1 vs Q2 rows and firms
  court_stays_by_buyer.png    - stayed/revoked/superseded rows by buyer
  headline_cases_timeline.png - Zee/Meril/Nestor: other buyer's ban
                                 period, RMSC bid deadline(s), RC period
"""

import textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from collections import Counter, defaultdict
from datetime import date

from openpyxl import load_workbook

from common import PROJECT
from analysis_common import (ANALYSIS_DIR, FROZEN, load_master_rows, is_active_ban, parse_date,
                              PALETTE, INK_PRIMARY, INK_SECONDARY, INK_MUTED, GRIDLINE, BASELINE,
                              SURFACE)

plt.rcParams.update({
    "font.family": "sans-serif", "font.size": 10.5, "text.color": INK_PRIMARY,
    "axes.edgecolor": BASELINE, "axes.labelcolor": INK_SECONDARY, "xtick.color": INK_MUTED,
    "ytick.color": INK_MUTED, "figure.facecolor": SURFACE, "axes.facecolor": SURFACE,
    "savefig.facecolor": SURFACE,
})

SOURCE_LINE = ("Source: output/frozen_2026-11-07/master_list.xlsx, blacklist_gap project "
               "(9 Indian government drug/surgical-goods buyers) - {extra}")


def _source(ax, extra, y=-0.16, width=100):
    text = SOURCE_LINE.format(extra=extra)
    wrapped = "\n".join(textwrap.wrap(text, width=width))
    ax.annotate(wrapped, xy=(0, y), xycoords="axes fraction",
                fontsize=8, color=INK_MUTED, ha="left", va="top")


def _style_axes(ax):
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color(BASELINE)
    ax.spines["bottom"].set_color(BASELINE)
    ax.grid(axis="y", color=GRIDLINE, linewidth=0.8, zorder=0)
    ax.set_axisbelow(True)


def chart_bans_by_buyer_reason(firm_rows, idx):
    BUYER, CAT = idx["Buyer"], idx["Reason category (Quality / Non-supply / Other / To check)"]
    buyers = sorted({r[BUYER] for r in firm_rows})
    cats = ["Quality", "Non-supply", "Other", "To check"]
    colors = [PALETTE["red"], PALETTE["blue"], PALETTE["violet"], PALETTE["yellow"]]
    counts = {cat: [sum(1 for r in firm_rows if r[BUYER] == b and r[CAT] == cat) for b in buyers]
              for cat in cats}

    fig, ax = plt.subplots(figsize=(8.5, 6))
    bottom = [0] * len(buyers)
    for cat, color in zip(cats, colors):
        ax.bar(buyers, counts[cat], bottom=bottom, color=color, label=cat,
               edgecolor=SURFACE, linewidth=1.5, zorder=2, width=0.68)
        bottom = [b + c for b, c in zip(bottom, counts[cat])]
    _style_axes(ax)
    ax.set_ylabel("Number of ban rows")
    ax.set_title("Bans per buyer, by reason category", fontsize=13, fontweight="bold",
                 color=INK_PRIMARY, pad=14)
    ax.legend(frameon=False, loc="upper right", fontsize=9.5)
    plt.xticks(rotation=20, ha="right")
    _source(ax, "every row with Level != Lab, all reason categories, not just active bans.",
           width=72)
    fig.tight_layout()
    fig.savefig(ANALYSIS_DIR / "bans_by_buyer_reason.png", dpi=150)
    plt.close(fig)


def chart_q1_vs_q2(firm_rows, idx):
    QLEVEL, FIRM = idx["Quality level (Q1 = statutory lab named, Q2 = other/no lab)"], idx["Firm name (standard)"]
    rows_by_level = Counter(r[QLEVEL] for r in firm_rows)
    firms_by_level = defaultdict(set)
    for r in firm_rows:
        firms_by_level[r[QLEVEL]].add(r[FIRM])

    levels = ["Q1", "Q2"]
    labels = ["Q1\n(named statutory lab)", "Q2\n(buyer's own/empanelled lab,\nor no lab named)"]
    row_vals = [rows_by_level.get(l, 0) for l in levels]
    firm_vals = [len(firms_by_level.get(l, set())) for l in levels]

    fig, ax = plt.subplots(figsize=(7.5, 6))
    x = range(len(levels))
    w = 0.32
    ax.bar([i - w / 2 for i in x], row_vals, width=w, color=PALETTE["blue"], label="Rows",
           edgecolor=SURFACE, linewidth=1.2, zorder=2)
    ax.bar([i + w / 2 for i in x], firm_vals, width=w, color=PALETTE["orange"],
           label="Distinct firms", edgecolor=SURFACE, linewidth=1.2, zorder=2)
    for i, (rv, fv) in enumerate(zip(row_vals, firm_vals)):
        ax.annotate(str(rv), (i - w / 2, rv), textcoords="offset points", xytext=(0, 4),
                    ha="center", fontsize=9.5, color=INK_SECONDARY)
        ax.annotate(str(fv), (i + w / 2, fv), textcoords="offset points", xytext=(0, 4),
                    ha="center", fontsize=9.5, color=INK_SECONDARY)
    _style_axes(ax)
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels)
    ax.set_ylabel("Count")
    ax.set_title("Quality bans: Q1 vs Q2", fontsize=13, fontweight="bold", color=INK_PRIMARY,
                 pad=14)
    ax.legend(frameon=False, loc="upper right")
    _source(ax, "every row with Level != Lab, Category = Quality or N/A (Q1/Q2 computed for "
                "Quality rows only).", width=60)
    fig.tight_layout()
    fig.savefig(ANALYSIS_DIR / "q1_vs_q2.png", dpi=150)
    plt.close(fig)


def chart_court_stays_by_buyer(firm_rows, idx):
    BUYER, LEVEL, COURT = idx["Buyer"], idx["Level (firm / product)"], idx["Court stay or revoked?"]
    excluded = [r for r in firm_rows if not is_active_ban(r[COURT], r[LEVEL])]
    by_buyer = Counter(r[BUYER] for r in excluded)
    buyers = sorted(by_buyer, key=lambda b: -by_buyer[b])
    vals = [by_buyer[b] for b in buyers]

    fig, ax = plt.subplots(figsize=(8, 6))
    ax.bar(buyers, vals, color=PALETTE["violet"], edgecolor=SURFACE, linewidth=1.2, zorder=2)
    for i, v in enumerate(vals):
        ax.annotate(str(v), (i, v), textcoords="offset points", xytext=(0, 4), ha="center",
                    fontsize=9.5, color=INK_SECONDARY)
    _style_axes(ax)
    ax.set_ylabel("Rows genuinely stayed, revoked, or superseded")
    ax.set_title("Court stays / revocations, by buyer", fontsize=13, fontweight="bold",
                 color=INK_PRIMARY, pad=14)
    plt.xticks(rotation=20, ha="right")
    _source(ax, "a genuine judicial stay/revocation/supersession only - excludes a plain "
                "'blacklisting period completed' note.", width=65)
    fig.tight_layout()
    fig.savefig(ANALYSIS_DIR / "court_stays_by_buyer.png", dpi=150)
    plt.close(fig)


# Confirmed RMSC bid deadlines - SAME constants as build_rc_vs_bans.py's
# TENDER_DEADLINES (not re-derived): NIB-10/2024 and NIB-14/2024 from
# native PDF text; NIB-01/2025 OCR'd, user-confirmed (checklist item 4,
# 2026-11-07); NIB-10/2026 OCR'd, user-confirmed by eye pages 1-3
# (checklist item 17, 2026-11-07).
DEADLINES = {
    "NIB-10/2024": date(2024, 10, 8), "NIB-14/2024": date(2025, 2, 24),
    "NIB-01/2025": date(2025, 6, 5), "NIB-10/2026": date(2026, 5, 6),
}

TIMELINE_CASES = [
    {"firm": "ZEE LABORATORIES LTD", "label": "Zee Laboratories\n(MPPHSCL ban, Madhya Pradesh)",
     "ban_buyer": "MPPHSCL", "deadlines": ["NIB-10/2024", "NIB-14/2024", "NIB-01/2025"],
     "rc_from": date(2025, 5, 31), "rc_to": date(2027, 9, 30)},
    {"firm": "MERIL ENDO SURGERY PVT LTD", "label": "Meril Endo Surgery\n(GMSCL ban, Gujarat)",
     "ban_buyer": "GMSCL", "deadlines": ["NIB-10/2026"],
     "rc_from": date(2026, 8, 20), "rc_to": date(2028, 7, 31)},
    {"firm": "NESTOR PHARMACEUTICALS LTD", "label": "Nestor Pharmaceuticals\n(APMSIDC ban, "
     "Andhra Pradesh)", "ban_buyer": "APMSIDC", "deadlines": ["NIB-14/2024"],
     "rc_from": date(2025, 9, 29), "rc_to": date(2027, 3, 31)},
]


def chart_headline_timeline(rows, idx):
    FIRM, BUYER, CAT = idx["Firm name (standard)"], idx["Buyer"], idx["Reason category (Quality / Non-supply / Other / To check)"]
    LEVEL, COURT = idx["Level (firm / product)"], idx["Court stay or revoked?"]
    FROM, TO = idx["Blacklisted from"], idx["Blacklisted to"]
    QLEVEL = idx["Quality level (Q1 = statutory lab named, Q2 = other/no lab)"]

    fig, ax = plt.subplots(figsize=(11, 5.5))
    y_positions = list(range(len(TIMELINE_CASES)))[::-1]

    for y, case in zip(y_positions, TIMELINE_CASES):
        earliest_deadline = min(DEADLINES[d] for d in case["deadlines"])
        # Only the CONFIRMED Q1 rows for this specific buyer, AND only those
        # still in force by the time of the firm's earliest relevant RMSC
        # deadline - e.g. Zee also carries older/unconfirmed (Q2) and a
        # long-since-expired Q1 MPPHSCL ban (Ranitidine, order 8078-2021,
        # expired 12-08-2023 - itself before every RMSC deadline here); both
        # kinds would otherwise wrongly stretch this bar back to 2020.
        ban_rows = [r for r in rows if r[FIRM] == case["firm"] and r[BUYER] == case["ban_buyer"]
                    and r[CAT] == "Quality" and r[QLEVEL] == "Q1"
                    and is_active_ban(r[COURT], r[LEVEL])]
        froms = [parse_date(r[FROM]) for r in ban_rows]
        tos = [parse_date(r[TO]) for r in ban_rows]
        pairs = [(f, t) for f, t in zip(froms, tos) if f and t and t >= earliest_deadline]
        ban_from, ban_to = min(p[0] for p in pairs), max(p[1] for p in pairs)

        ax.barh(y + 0.18, (ban_to - ban_from).days, left=ban_from, height=0.28,
                color=PALETTE["red"], zorder=3, label="Other buyer's ban period (confirmed Q1 "
                "rows only)" if y == y_positions[0] else None)
        ax.barh(y - 0.18, (case["rc_to"] - case["rc_from"]).days, left=case["rc_from"],
                height=0.28, color=PALETTE["blue"], zorder=3,
                label="RMSC rate-contract period" if y == y_positions[0] else None)
        for i, dl_code in enumerate(case["deadlines"]):
            dl = DEADLINES[dl_code]
            ax.plot(dl, y, marker="v", markersize=12, color=PALETTE["orange"],
                    markeredgecolor=SURFACE, markeredgewidth=1, zorder=4,
                    label="RMSC bid deadline" if (y == y_positions[0] and i == 0) else None)
            ax.annotate(dl_code, (dl, y), textcoords="offset points", xytext=(0, 8 + 10 * i),
                        ha="center", fontsize=7.5, color=INK_SECONDARY)

    ax.set_yticks(y_positions)
    ax.set_yticklabels([c["label"] for c in TIMELINE_CASES], fontsize=10)
    ax.set_ylim(min(y_positions) - 0.6, max(y_positions) + 1.0)
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.set_xlim(date(2021, 1, 1), date(2029, 1, 1))
    _style_axes(ax)
    ax.grid(axis="x", color=GRIDLINE, linewidth=0.8, zorder=0)
    ax.set_xlabel("Year")
    ax.set_title("The 3 headline cases: other buyer's ban vs Rajasthan tender deadline vs RC period",
                 fontsize=12.5, fontweight="bold", color=INK_PRIMARY, pad=16)
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles, labels, frameon=False, loc="upper left", bbox_to_anchor=(0, -0.20), ncol=3,
             fontsize=9.5)
    _source(ax, "ban period = active, Quality-category rows for the named buyer only; RC period "
                "= min/max Contract From/To across the firm's clause-(ii)-scope RMSC lines; "
                "deadlines are this project's own confirmed figures (see "
                "output/finance/headline_case_values.xlsx).", y=-0.42, width=100)
    fig.subplots_adjust(bottom=0.44, top=0.86)
    fig.savefig(ANALYSIS_DIR / "headline_cases_timeline.png", dpi=150)
    plt.close(fig)


def main():
    idx, rows = load_master_rows()
    firm_rows = [r for r in rows if r[idx["Level (firm / product)"]] != "Lab"]

    chart_bans_by_buyer_reason(firm_rows, idx)
    chart_q1_vs_q2(firm_rows, idx)
    chart_court_stays_by_buyer(firm_rows, idx)
    chart_headline_timeline(rows, idx)

    print("Saved 4 charts to", ANALYSIS_DIR)
    for f in ["bans_by_buyer_reason.png", "q1_vs_q2.png", "court_stays_by_buyer.png",
              "headline_cases_timeline.png"]:
        print(" ", ANALYSIS_DIR / f)


if __name__ == "__main__":
    main()
