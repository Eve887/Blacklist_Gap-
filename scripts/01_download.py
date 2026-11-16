"""
01_download.py - STEP 2a: download the blacklist PDFs.

What it does, in order:
  1. Downloads the two Andhra Pradesh (APMSIDC) PDFs.
  2. Opens the Assam (AMSCL) blacklist web page, finds the 12 rows,
     and downloads the notice PDF linked in each row.
  3. Writes a log of every download (output/download_log.csv) with a
     "fingerprint" (SHA-256) of each file. The fingerprint proves later
     that the file was not changed after you downloaded it.

Safe to run again: files you already have are skipped.
Safe to stop with Ctrl+C: the log is saved before the script exits.
"""

import csv
import re
import hashlib            # makes the SHA-256 fingerprint of a file
import time               # lets the script pause between downloads
from datetime import datetime
from urllib.parse import urljoin, urlparse, unquote
from urllib import robotparser   # reads a website's robots.txt rules

import requests                  # downloads web pages and files
from bs4 import BeautifulSoup    # reads the HTML of a web page

from common import RAW, OUT, one_line

# ------------------------------------------------------------------
# SETTINGS - change these if you need to
# ------------------------------------------------------------------
AP_FILES = [
    # (buyer folder, file name to save as, web address)
    ("AP_APMSIDC", "AP_firms_blacklisted_2022-12-13.pdf",
     "https://apmsidc.ap.nic.in/ftp/drugs/Firm_Blacklisted_APMSIDC_UpdatedDt 13-12-2022.pdf"),
    ("AP_APMSIDC", "AP_items_blacklisted_2024-09-13.pdf",
     "https://apmsidc.ap.nic.in/ftp/drugs/Items_Blacklisted_APMSIDC_UpdatedDt 13-09-2024 (1).pdf"),
    # Older versions of the SAME items list, kept as SEPARATE files (never overwrite
    # the 2024 one) so later versions can be compared against earlier ones.
    ("AP_APMSIDC", "AP_items_blacklisted_2022-12-13.pdf",
     "https://apmsidc.ap.nic.in/ftp/drugs/Items_Blacklisted_APMSIDC_UpdatedDt 13-12-2022.pdf"),
    ("AP_APMSIDC", "AP_items_blacklisted_2023-11-30.pdf",
     "https://apmsidc.ap.nic.in/ftp/drugs/Items_Blacklisted_APMSIDC_UpdatedDt 30-11-2023.pdf"),
]
ASSAM_PAGE = "https://amscl.assam.gov.in/information-and-services/blacklisted-debarred-firms"

WAIT_SECONDS = 3      # pause between downloads, so we do not overload government servers
TRIES = 3             # how many times to retry a failed download
TIMEOUT = 60          # give up on one attempt after 60 seconds
HEADERS = {"User-Agent": "Mozilla/5.0 (research project; blacklist-gap study)"}

LOG_FILE = OUT / "download_log.csv"
ASSAM_INDEX = OUT / "assam_index.csv"

session = requests.Session()     # one Session re-uses the same connection = faster
session.headers.update(HEADERS)
log_rows = []                    # every download result is collected here


# ------------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------------
_robots_cache = {}

def robots_allows(url):
    """Ask the website's robots.txt whether automatic downloading is allowed.
    If the site forbids it, we do NOT download - we tell you to use your browser."""
    root = "{0.scheme}://{0.netloc}".format(urlparse(url))
    if root not in _robots_cache:
        rp = robotparser.RobotFileParser()
        try:
            r = session.get(root + "/robots.txt", timeout=20)
            if r.status_code == 200:
                rp.parse(r.text.splitlines())
            else:
                rp.parse([])          # no robots.txt = no rules = allowed
        except requests.RequestException:
            rp.parse([])
        _robots_cache[root] = rp
    return _robots_cache[root].can_fetch(HEADERS["User-Agent"], url)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def download(url, folder, filename):
    """Download one file. Returns a short status word."""
    target = RAW / folder / filename
    target.parent.mkdir(parents=True, exist_ok=True)

    if target.exists() and target.stat().st_size > 0:
        print(f"  already have  {filename}")
        return record(url, target, "already had")

    if not robots_allows(url):
        print(f"  BLOCKED by robots.txt -> download it in your browser and save it as:\n"
              f"      {target}\n      link: {url}")
        return record(url, target, "blocked by robots.txt - download manually")

    for attempt in range(1, TRIES + 1):
        try:
            r = session.get(url, timeout=TIMEOUT)
            r.raise_for_status()                      # error if the site said 404, 500 etc.
            if not r.content.startswith(b"%PDF"):     # every real PDF begins with "%PDF"
                print(f"  NOT A PDF (the site sent a web page instead): {url}")
                return record(url, target, "not a pdf")
            target.write_bytes(r.content)
            print(f"  downloaded    {filename}  ({len(r.content)//1024} KB)")
            time.sleep(WAIT_SECONDS)
            return record(url, target, "downloaded")
        except requests.RequestException as e:
            print(f"  attempt {attempt}/{TRIES} failed: {e}")
            time.sleep(WAIT_SECONDS * attempt)        # wait longer after each failure
    return record(url, target, "failed")


def record(url, path, status):
    log_rows.append({
        "date_time": datetime.now().strftime("%d-%m-%Y %H:%M"),
        "status": status,
        "file": str(path.relative_to(RAW.parent.parent)),
        "sha256": sha256(path) if path.exists() else "",
        "url": url,
    })
    return status


def save_log():
    with open(LOG_FILE, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig = opens cleanly in Excel
        w = csv.DictWriter(f, fieldnames=["date_time", "status", "file", "sha256", "url"])
        w.writeheader()
        w.writerows(log_rows)
    print(f"\nLog saved: {LOG_FILE}")


def safe_filename(text):
    keep = "".join(c if c.isalnum() or c in " -_" else "_" for c in text)
    return "_".join(keep.split())[:80]


# ------------------------------------------------------------------
# PART 1 - Andhra Pradesh
# ------------------------------------------------------------------
def andhra_pradesh():
    print("\n=== Andhra Pradesh (APMSIDC) ===")
    for folder, name, url in AP_FILES:
        download(url, folder, name)


# ------------------------------------------------------------------
# PART 2 - Assam: read the web page, then download each notice
# ------------------------------------------------------------------
def assam():
    print("\n=== Assam (AMSCL) ===")
    if not robots_allows(ASSAM_PAGE):
        print("  Assam page is blocked by robots.txt - download the notices in your browser.")
        return
    r = session.get(ASSAM_PAGE, timeout=TIMEOUT)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    index = []
    for tr in soup.find_all("tr"):                     # every table row on the page
        cells = tr.find_all(["td", "th"])
        link = tr.find("a", href=True)
        if len(cells) < 2 or not link:
            continue                                  # skip header rows and empty rows
        texts = [one_line(c.get_text(" ")) for c in cells]
        sl_no = texts[0]
        name_cell = texts[1]                          # e.g. "Mercury Laboratories Limited Date: 31-03-2023"
        firm, _, date = name_cell.partition("Date:")
        if not date:                                  # date may sit in a different cell
            date = next((t for t in texts if re.search(r"\d{2}-\d{2}-\d{4}", t)), "")
            date = date.replace("Date:", "")
        pdf_url = urljoin(ASSAM_PAGE, link["href"])   # turns "/sites/..." into a full web address
        fname = f"AS_{int(sl_no):02d}_{safe_filename(firm)}.pdf" if sl_no.isdigit() else \
                safe_filename(unquote(pdf_url.rsplit('/', 1)[-1])) + ".pdf"
        status = download(pdf_url, "AS_AMSCL", fname)
        index.append({"sl_no": sl_no, "firm_as_written": firm.strip(), "date": date.strip(),
                      "file": fname, "pdf_url": pdf_url, "status": status})

    if not index:
        print("  No rows found. The page layout may have changed - open it in your browser.")
    with open(ASSAM_INDEX, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["sl_no", "firm_as_written", "date", "file", "pdf_url", "status"])
        w.writeheader()
        w.writerows(index)
    print(f"  {len(index)} rows found. Index saved: {ASSAM_INDEX}")


# ------------------------------------------------------------------
# RUN
# ------------------------------------------------------------------
if __name__ == "__main__":
    try:
        andhra_pradesh()
        assam()
    except KeyboardInterrupt:
        print("\nStopped by you (Ctrl+C). Saving what was done so far...")
    finally:
        save_log()      # "finally" runs even if the script was stopped or crashed
