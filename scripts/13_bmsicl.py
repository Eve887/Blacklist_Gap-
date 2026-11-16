"""
11_bmsicl.py - Bihar Medical Services & Infrastructure Corporation Limited
(BMSICL) buyer file.

BMSICL publishes each blacklisting/debarment action as its own WordPress
post (title = the firm's name), not a table. The post pages themselves are
nearly empty (just a link to a PDF order letter) - no visible date or
reason text in the HTML. Two things make this workable without scraping
~9 pages of the HTML archive and ~80 individual post pages:
  - The public RSS feed (".../feed/") gives title + publish date + the
    order-PDF link for every post, paginated with ?paged=N. This is a
    normal public syndication feed, not an access-restricted API - the
    site's own REST API (wp-json) IS restricted (401, confirmed) and was
    NOT used.
  - Each post carries its own WordPress post-id (visible in the feed's
    embedded HTML), which is a stable, source-given identifier - used
    directly for Entry IDs instead of trying to parse an inconsistent
    order number out of free-text link titles.

The order PDFs themselves are ALSO scanned images (confirmed: 0 extractable
characters) - same OCR blocker as KSMSCL (see source_issues.csv). So
"Blacklisted from/to" and "Reason (as written)" are honestly "Not stated"
here, not guessed. The WordPress PUBLISH date is kept as supplementary
context in "Page / row in source" (labelled as a publish date, not
presented as the legal order's effective date).

Revocations are posted in the SAME feed/category as debarments (e.g.
"Revoking of blacklisting order of M/s Kanha Biogenetic") - matched
against this buyer's own debarment rows by standardized firm name; a
revocation that cannot be matched goes to "Unmatched revocations" instead
of being dropped, same pattern as PMBI.

Writes:
  output/buyers/BMSICL.xlsx  with two sheets:
    - Master list           : BMSICL rows, in the 19 master-list columns
    - Unmatched revocations : revocations that could not be matched
"""

import re
import time
from datetime import date, datetime
from email.utils import parsedate_to_datetime
from urllib.parse import urlparse
from urllib import robotparser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

from common import BUYERS, COLUMNS, standard_name, one_line, dedup_entry_ids, write_sheet, CATEGORY_PENDING

BASE = "http://bmsicl.gov.in/new/"
FEED = "http://bmsicl.gov.in/new/drug-tender/drug-blacklist-debarred/feed/"

HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}
WAIT_SECONDS = 3
TIMEOUT = 60
TODAY = date.today().strftime("%d-%m-%Y")
MAX_PAGES = 30

session = requests.Session()
session.headers.update(HEADERS)

_robots_cache = {}


def robots_allows(url):
    root = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if root not in _robots_cache:
        rp = robotparser.RobotFileParser()
        try:
            r = session.get(root + "/robots.txt", timeout=20)
            if r.status_code == 200:
                rp.parse(r.text.splitlines())
            else:
                rp.parse([])
        except requests.RequestException:
            rp.parse([])
        _robots_cache[root] = rp
    return _robots_cache[root].can_fetch(HEADERS["User-Agent"], url)


LINK_RE = re.compile(r'<a\s+href="([^"]+download\.php[^"]+)"[^>]*post-id="(\d+)"[^>]*>([^<]*)</a>')


def fetch_feed_items():
    items = []
    for page_no in range(1, MAX_PAGES + 1):
        url = FEED if page_no == 1 else f"{FEED}?paged={page_no}"
        if not robots_allows(url):
            print(f"  BLOCKED by robots.txt - open in a browser: {url}")
            break
        r = session.get(url, timeout=TIMEOUT)
        time.sleep(WAIT_SECONDS)
        if r.status_code == 404:
            break
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "xml")
        page_items = soup.find_all("item")
        if not page_items:
            break
        items.extend(page_items)
        print(f"  feed page {page_no}: {len(page_items)} items")
    return items


def parse_item(item):
    title = one_line(item.find("title").get_text())
    link = item.find("link").get_text()
    pub_raw = item.find("pubDate").get_text() if item.find("pubDate") else ""
    try:
        pub_dt = parsedate_to_datetime(pub_raw)
        pub_str = pub_dt.strftime("%d-%m-%Y")
    except (TypeError, ValueError):
        pub_str = "Not stated"
    content = item.find("encoded")
    content_html = content.get_text() if content is not None else ""
    m = LINK_RE.search(content_html)
    if m:
        doc_url, post_id, link_text = m.group(1), m.group(2), m.group(3).strip()
    else:
        doc_url, post_id, link_text = link, None, ""
    return {"title": title, "link": link, "pub": pub_str, "doc_url": doc_url,
            "post_id": post_id, "link_text": link_text}


REVOKE_PREFIX_RE = re.compile(r"^revoking\s+of\s+blacklisting\s+order\s+of\s+", re.I)


def main():
    raw_items = fetch_feed_items()
    parsed = [parse_item(it) for it in raw_items]

    debar_rows = []
    debar_by_name = {}
    revoke_items = []
    for it in parsed:
        if REVOKE_PREFIX_RE.match(it["title"]):
            revoke_items.append(it)
            continue
        firm = it["title"].rstrip(".").strip()
        entry_id = f"BMSICL-DRUG-{it['post_id']}" if it["post_id"] else f"BMSICL-DRUG-{it['link']}"
        where = (f"WordPress post {it['post_id']}, published {it['pub']}, "
                 f"doc title \"{it['link_text']}\" ({it['link']})")
        row = [entry_id, one_line(firm), standard_name(firm), "Not stated", "BMSICL", "Bihar",
               "Firm (inferred)", "-", "Not stated", "Not stated", "Not stated", "Not stated",
               CATEGORY_PENDING, "Not stated", "Not stated", it["doc_url"], where, "Auto - check", TODAY]
        debar_rows.append(row)
        debar_by_name.setdefault(standard_name(firm), []).append(row)

    unmatched = []
    for it in revoke_items:
        firm = REVOKE_PREFIX_RE.sub("", it["title"]).rstrip(".").strip()
        target = standard_name(firm)
        matches = debar_by_name.get(target)
        if matches:
            note = f"Revoked - published {it['pub']}"
            for r in matches:
                r[14] = note
        else:
            unmatched.append(["BMSICL", one_line(firm), target, "Not stated", it["pub"], it["doc_url"],
                               f"Could not match to any BMSICL debarment post by standardized name "
                               f"(checked {len(debar_rows)} debarment posts from this buyer only). "
                               f"Revocation post title: \"{it['title']}\"."])

    dupe_warnings = dedup_entry_ids(debar_rows)
    if dupe_warnings:
        print("WARNING - duplicate Entry IDs found, '-b'/'-c' added to tell them apart:")
        for w in dupe_warnings:
            print("  ", w)

    wb = Workbook()
    write_sheet(wb.active, COLUMNS, debar_rows, [20, 34, 30, 14, 10, 10, 16, 10, 13, 13, 12, 14, 14, 14, 20, 50, 60, 16, 12])
    wb.active.title = "Master list"
    write_sheet(wb.create_sheet("Unmatched revocations"),
                ["Buyer", "Firm name (as written)", "Firm name (standard)", "Reason (as written)",
                 "Revocation date", "Source link", "Note"],
                unmatched, [10, 34, 30, 14, 16, 50, 70])
    out = BUYERS / "BMSICL.xlsx"
    wb.save(out)

    print(f"Saved: {out}")
    print(f"  Master list rows      : {len(debar_rows)}")
    print(f"  Revocation posts seen : {len(revoke_items)}")
    print(f"  Unmatched revocations : {len(unmatched)}")


if __name__ == "__main__":
    main()
