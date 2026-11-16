"""
build_write_up_exports.py - FINAL STAGE Part 2. Converts the write_up/
markdown deliverables into the extra formats asked for:
  - Blacklist_Gap_Report.md  -> Blacklist_Gap_Report.docx (python-docx)
  - One_page_summary.md      -> One_page_summary.pdf (matplotlib, one page,
                                 text + the headline-cases timeline chart)

No pandoc/LibreOffice/reportlab is installed in this environment (checked
before writing this script) - the .docx is built directly with python-docx
against a small, constrained Markdown subset (this project's own write-up
files only use: #/##/### headings, **bold**, *italic*, tables, blockquotes,
horizontal rules, bullet/numbered lists, and [text](url) links) rather than
a general-purpose Markdown parser. The .pdf is composed directly as a
single matplotlib page (text blocks + one embedded PNG), since there is no
docx/md -> pdf converter available either.

Not part of the numbered buyer pipeline - a one-off formatting step, run
after the .md files themselves are finished and not expected to change
independently of them.
"""

import re
import textwrap
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.image as mpimg

from common import PROJECT

WRITE_UP = PROJECT / "write_up"
ANALYSIS = PROJECT / "output" / "analysis"

LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
ITALIC_RE = re.compile(r"(?<!\*)\*([^*]+?)\*(?!\*)")


def _clean_inline(text):
    """Strips markdown link syntax to 'text (url)' and returns plain text
    with bold/italic markers removed (bold/italic handled separately by
    _add_runs for docx; this version is for the PDF path, which does not
    need rich runs)."""
    text = LINK_RE.sub(r"\1 (\2)", text)
    text = BOLD_RE.sub(r"\1", text)
    text = ITALIC_RE.sub(r"\1", text)
    return text


def _add_runs(paragraph, text):
    """Adds bold/italic runs to a docx paragraph from a constrained
    Markdown inline subset (**bold**, *italic*, [text](url) -> 'text
    (url)') - a small hand-written tokenizer, not a general parser."""
    text = LINK_RE.sub(r"\1 (\2)", text)
    pos = 0
    token_re = re.compile(r"\*\*(.+?)\*\*|(?<!\*)\*([^*]+?)\*(?!\*)")
    for m in token_re.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        if m.group(1) is not None:
            r = paragraph.add_run(m.group(1))
            r.bold = True
        else:
            r = paragraph.add_run(m.group(2))
            r.italic = True
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def _join_paragraphs(lines):
    """Standard Markdown treats a run of non-blank, non-block-starting
    lines as ONE paragraph/list-item/blockquote - this project's own .md
    files wrap long paragraphs AND long list items across several source
    lines. Processing `lines` one at a time (as a first version of this
    script did) can split a **bold** or *italic* marker pair across 2
    source lines, so the closing marker is never seen in the same regex
    pass and nothing gets stripped - found exactly this bug twice while
    building this script (once in a plain paragraph, once in a wrapped
    numbered-list item) before writing this proper version. Returns a new
    list where each wrapped paragraph/list-item/blockquote becomes ONE
    joined string (continuation lines appended with a space, the item's
    own leading "- "/"N. "/"> " marker kept on the first chunk); table
    rows and headings are never joined (each stays its own source line,
    since this project's own files never wrap those) and a blank line
    simply ends whatever block was open."""
    out, buf, buf_type = [], [], None

    def flush():
        nonlocal buf_type
        if buf:
            out.append(" ".join(buf))
            buf.clear()
        buf_type = None

    for line in lines:
        stripped = line.strip()
        if not stripped:
            flush()
        elif stripped == "---" or stripped.startswith("#") or stripped.startswith("|") or \
                (stripped.startswith("*") and stripped.endswith("*") and len(stripped) > 1):
            flush()
            out.append(line)
        elif stripped.startswith("- ") or stripped.startswith("* ") or \
                re.match(r"^\d+\.\s", stripped):
            flush()
            buf.append(stripped)
            buf_type = "list"
        elif stripped.startswith("> "):
            # This project's own blockquotes repeat "> " on EVERY wrapped
            # line (unlike its list items, which only mark the first
            # line) - a continuation line re-matching this same branch
            # must be appended to the open quote block, not treated as
            # the start of a new one (that bug split the clause-ii quote
            # into 11 separate one-line "quotes" before this fix). The
            # leading "> " is kept ONLY on the first chunk, so the joined
            # result still starts with "> " for the branches below that
            # test for it, but the marker isn't repeated mid-sentence.
            if buf_type != "quote":
                flush()
                buf_type = "quote"
                buf.append(stripped)
            else:
                buf.append(stripped[2:])
        else:
            if buf_type is None:
                buf_type = "para"
            buf.append(stripped)
    flush()
    return out


def _parse_table(lines, i):
    """Reads a markdown pipe-table starting at lines[i] (the header row).
    Returns (rows, next_i)."""
    rows = []
    header = [c.strip() for c in lines[i].strip().strip("|").split("|")]
    rows.append(header)
    i += 2                                            # skip the |---|---| separator row
    while i < len(lines) and lines[i].strip().startswith("|"):
        rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
        i += 1
    return rows, i


def md_to_docx(md_path, docx_path, title):
    lines = _join_paragraphs(Path(md_path).read_text(encoding="utf-8").splitlines())
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue
        if stripped == "---":
            doc.add_paragraph().add_run("_" * 60).font.color.rgb = RGBColor(0xC3, 0xC2, 0xB7)
            i += 1
            continue
        if stripped.startswith("#### "):
            doc.add_heading(_clean_inline(stripped[5:]), level=4)
            i += 1
            continue
        if stripped.startswith("### "):
            doc.add_heading(_clean_inline(stripped[4:]), level=3)
            i += 1
            continue
        if stripped.startswith("## "):
            doc.add_heading(_clean_inline(stripped[3:]), level=2)
            i += 1
            continue
        if stripped.startswith("# "):
            h = doc.add_heading(_clean_inline(stripped[2:]), level=1)
            i += 1
            continue
        if stripped.startswith("*") and stripped.endswith("*") and not stripped.startswith("**"):
            p = doc.add_paragraph()
            r = p.add_run(_clean_inline(stripped.strip("*")))
            r.italic = True
            i += 1
            continue
        if stripped.startswith("> "):
            p = doc.add_paragraph(style="Intense Quote")
            _add_runs(p, stripped[2:])
            i += 1
            continue
        if stripped.startswith("|"):
            rows, i = _parse_table(lines, i)
            table = doc.add_table(rows=len(rows), cols=len(rows[0]))
            table.style = "Light Grid Accent 1"
            for r_i, row in enumerate(rows):
                for c_i, cell_text in enumerate(row):
                    if c_i >= len(table.columns):
                        continue
                    cell = table.cell(r_i, c_i)
                    cell.text = ""
                    p = cell.paragraphs[0]
                    _add_runs(p, cell_text)
                    if r_i == 0:
                        for run in p.runs:
                            run.bold = True
            doc.add_paragraph()
            continue
        if re.match(r"^\d+\.\s", stripped):
            p = doc.add_paragraph(style="List Number")
            _add_runs(p, re.sub(r"^\d+\.\s", "", stripped))
            i += 1
            continue
        if stripped.startswith("- ") or stripped.startswith("* "):
            p = doc.add_paragraph(style="List Bullet")
            _add_runs(p, stripped[2:])
            i += 1
            continue

        p = doc.add_paragraph()
        _add_runs(p, stripped)
        i += 1

    doc.core_properties.title = title
    doc.save(docx_path)
    print(f"Saved: {docx_path}")


def one_pager_to_pdf(md_path, pdf_path, chart_path):
    joined_lines = _join_paragraphs(Path(md_path).read_text(encoding="utf-8").splitlines())
    fig = plt.figure(figsize=(8.5, 11))
    fig.patch.set_facecolor("#fcfcfb")
    ax = fig.add_axes([0.07, 0.30, 0.86, 0.66])
    ax.axis("off")

    y = 1.0
    for raw_line in joined_lines:
        line = raw_line.rstrip()
        if not line:
            y -= 0.012
            continue
        if line.startswith("# "):
            ax.text(0, y, _clean_inline(line[2:]), fontsize=16, fontweight="bold",
                    color="#0b0b0b", va="top")
            y -= 0.040
        elif line.startswith("## "):
            ax.text(0, y, _clean_inline(line[3:]), fontsize=12.5, fontweight="bold",
                    color="#0b0b0b", va="top")
            y -= 0.030
        elif line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if set("".join(cells)) <= set("-: "):
                continue                                # the |---|---| separator row
            row_text = "   ".join(_clean_inline(c) for c in cells)
            ax.text(0.01, y, textwrap.shorten(row_text, width=108), fontsize=7.6,
                    color="#0b0b0b", va="top", family="monospace")
            y -= 0.017
        elif line.startswith("- "):
            wrapped = textwrap.wrap(_clean_inline(line[2:]), width=104)
            for j, wl in enumerate(wrapped):
                ax.text(0.015, y, ("- " if j == 0 else "  ") + wl, fontsize=8.3,
                        color="#0b0b0b", va="top")
                y -= 0.0155
        elif line.startswith("*") and line.endswith("*"):
            wrapped = textwrap.wrap(_clean_inline(line.strip("*")), width=118)
            for wl in wrapped:
                ax.text(0, y, wl, fontsize=7.6, style="italic", color="#52514e", va="top")
                y -= 0.0145
        else:
            wrapped = textwrap.wrap(_clean_inline(line), width=108)
            for wl in wrapped:
                ax.text(0, y, wl, fontsize=8.3, color="#0b0b0b", va="top")
                y -= 0.0155
        if y < 0.02:
            break

    img_ax = fig.add_axes([0.07, 0.015, 0.86, 0.26])
    img_ax.axis("off")
    if Path(chart_path).exists():
        img = mpimg.imread(chart_path)
        img_ax.imshow(img, aspect="equal")

    fig.savefig(pdf_path, format="pdf", facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"Saved: {pdf_path}")


def main():
    md_to_docx(WRITE_UP / "Blacklist_Gap_Report.md", WRITE_UP / "Blacklist_Gap_Report.docx",
              "The Blacklist Gap")
    one_pager_to_pdf(WRITE_UP / "One_page_summary.md", WRITE_UP / "One_page_summary.pdf",
                     ANALYSIS / "headline_cases_timeline.png")


if __name__ == "__main__":
    main()
