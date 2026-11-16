"""
17_network_chart.py - Part 1 (data science), FINAL STAGE. READ-ONLY
REPORT: a bipartite firm<->buyer network chart - firms on the left,
buyers on the right, one line per ACTIVE ban (see
analysis_common.is_active_ban - excludes genuine court stays/
revocations/supersessions, keeps a plain "period completed" row). Only
firms with an active ban from 2 OR MORE distinct buyers are drawn, so
the chart stays readable (63 of 369 banned firms qualify - see
output/analysis/summary_numbers.xlsx "Firms by buyer count"). The 3
headline cases (Zee, Meril, Nestor) are highlighted in red; every other
edge/firm label is in a muted categorical color.

No networkx in this environment - laid out by hand (2 columns, firms
sorted by buyer count then name, buyers in a fixed order) rather than a
force-directed layout, which keeps the result reproducible run to run.

Writes output/analysis/network.png.
"""

import textwrap
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from collections import defaultdict

from analysis_common import (ANALYSIS_DIR, load_master_rows, is_active_ban, HEADLINE_FIRMS,
                              PALETTE, INK_PRIMARY, INK_SECONDARY, INK_MUTED, SURFACE)

MIN_BUYERS = 2
EDGE_COLOR_DEFAULT = "#c3c2b7"              # BASELINE - a recessive, non-categorical gray
EDGE_COLOR_HIGHLIGHT = PALETTE["red"]
NODE_COLOR_FIRM = PALETTE["blue"]
NODE_COLOR_BUYER = PALETTE["orange"]


def main():
    idx, rows = load_master_rows()
    FIRM, BUYER, LEVEL = idx["Firm name (standard)"], idx["Buyer"], idx["Level (firm / product)"]
    COURT = idx["Court stay or revoked?"]

    buyers_by_firm = defaultdict(set)
    for r in rows:
        if is_active_ban(r[COURT], r[LEVEL]):
            buyers_by_firm[r[FIRM]].add(r[BUYER])

    firms = sorted((f for f, b in buyers_by_firm.items() if len(b) >= MIN_BUYERS),
                   key=lambda f: (-len(buyers_by_firm[f]), f))
    buyers = sorted({b for f in firms for b in buyers_by_firm[f]})

    print(f"{len(firms)} firms with active bans from {MIN_BUYERS}+ distinct buyers "
          f"(of {len(buyers_by_firm)} banned firms total)")
    print(f"{len(buyers)} distinct buyers appear among these firms: {buyers}")

    n_firms, n_buyers = len(firms), len(buyers)
    fig_h = max(8, n_firms * 0.24)
    fig, ax = plt.subplots(figsize=(14, fig_h))

    firm_y = {f: n_firms - 1 - i for i, f in enumerate(firms)}
    buyer_y = {b: (n_firms - 1) * (i / max(1, n_buyers - 1)) for i, b in enumerate(buyers)}
    x_firm, x_buyer = 0.0, 1.0

    for f in firms:
        is_headline = f in HEADLINE_FIRMS
        for b in buyers_by_firm[f]:
            color = EDGE_COLOR_HIGHLIGHT if is_headline else EDGE_COLOR_DEFAULT
            lw = 1.8 if is_headline else 0.7
            alpha = 0.9 if is_headline else 0.45
            zorder = 3 if is_headline else 1
            ax.plot([x_firm, x_buyer], [firm_y[f], buyer_y[b]], color=color, linewidth=lw,
                    alpha=alpha, zorder=zorder, solid_capstyle="round")

    for f in firms:
        is_headline = f in HEADLINE_FIRMS
        ax.scatter([x_firm], [firm_y[f]], s=46 if is_headline else 22,
                  color=EDGE_COLOR_HIGHLIGHT if is_headline else NODE_COLOR_FIRM,
                  zorder=4, edgecolors=SURFACE, linewidths=0.6)
        label = f.title()
        ax.annotate(label, (x_firm, firm_y[f]), xytext=(-8, 0), textcoords="offset points",
                    ha="right", va="center",
                    fontsize=8.2 if not is_headline else 9.5,
                    fontweight="bold" if is_headline else "normal",
                    color=EDGE_COLOR_HIGHLIGHT if is_headline else INK_SECONDARY)

    for b in buyers:
        ax.scatter([x_buyer], [buyer_y[b]], s=120, color=NODE_COLOR_BUYER, zorder=4,
                  edgecolors=SURFACE, linewidths=0.8, marker="s")
        ax.annotate(b, (x_buyer, buyer_y[b]), xytext=(10, 0), textcoords="offset points",
                    ha="left", va="center", fontsize=11, fontweight="bold", color=INK_PRIMARY)

    ax.set_xlim(-0.55, 1.35)
    ax.set_ylim(-2.2, n_firms + 0.5)
    ax.axis("off")
    ax.set_title(f"Firms banned by {MIN_BUYERS}+ Indian government buyers\n"
                 f"({len(firms)} of {len(buyers_by_firm)} banned firms; Zee, Meril, Nestor "
                 f"highlighted in red)",
                 fontsize=14, fontweight="bold", color=INK_PRIMARY, pad=18, x=0.32)
    source = ("Source: output/frozen_2026-11-07/master_list.xlsx, blacklist_gap project - one "
              "line per ACTIVE ban (excludes genuine court stays/revocations/supersessions); "
              "firms with an active ban from only 1 buyer are not shown.")
    ax.annotate("\n".join(textwrap.wrap(source, width=110)), xy=(-0.55, -1.5), xycoords="data",
                fontsize=8.5, color=INK_MUTED, ha="left", va="top")

    # Manual margins, not tight_layout() - the long firm-name labels extend
    # well past the data xlim, and tight_layout()'s own bbox estimate
    # clipped both the title and the left-hand labels against the figure
    # edge when first tried.
    fig.subplots_adjust(left=0.22, right=0.97, top=0.93, bottom=0.03)
    out = ANALYSIS_DIR / "network.png"
    fig.savefig(out, dpi=150)
    plt.close(fig)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
