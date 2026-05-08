"""
NiftyArchive — BSE Annual Report Scraper
==========================================
Downloads annual reports (2015/16 - 2024/25) for Nifty 500 companies.

THE FIX: Instead of doing a live API lookup per company (which was returning
ABB India for everything), we download BSE's full scrip master list once and
do fuzzy name matching locally. This is reliable and fast.

Requirements:
    pip install requests openpyxl pandas bsedata

Usage:
    python bse_annual_report_scraper.py
"""

import requests
import pandas as pd
import os
import re
import time
import csv
import logging
from pathlib import Path
from difflib import get_close_matches
from datetime import datetime

# ─────────────────────────── CONFIGURATION ───────────────────────────────────

EXCEL_FILE  = "ind_nifty500list.xlsx"
OUTPUT_DIR  = "./annual_reports"
LOG_FILE    = "./scraper_log.csv"

YEARS = [
    "2015-16", "2016-17", "2017-18", "2018-19", "2019-20",
    "2020-21", "2021-22", "2022-23", "2023-24", "2024-25",
]

DELAY_BETWEEN_COMPANIES = 2    # seconds — be polite to BSE
DELAY_BETWEEN_REQUESTS  = 0.5
REQUEST_TIMEOUT         = 25

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer":         "https://www.bseindia.com/",
    "Accept":          "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
}

# ─────────────────────────── LOGGING ─────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("scraper.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ─────────────────────────── SCRIP CODE MASTER ───────────────────────────────

def build_scrip_master() -> dict:
    """
    Download BSE's full scrip master (all listed equities) and return a
    dict of {normalised_company_name: scrip_code}.

    Uses bsedata's updateScripCodes() which fetches the official BSE CSV.
    Falls back to a direct CSV download if bsedata is not installed.
    """
    # Method A: bsedata library
    try:
        from bsedata.bse import BSE as BSEData
        log.info("Building scrip master via bsedata ...")
        b = BSEData(update_codes=True)
        raw = b.getScripCodes()          # {scrip_code: company_name}
        master = {normalise(v): str(k) for k, v in raw.items()}
        log.info(f"  Scrip master loaded: {len(master):,} companies")
        return master
    except ImportError:
        pass
    except Exception as e:
        log.warning(f"bsedata failed: {e}")

    # Method B: direct BSE API
    log.info("Building scrip master via BSE API ...")
    url = "https://api.bseindia.com/BseIndiaAPI/api/ListofScripData/w"
    params = {"Group": "", "Scripcode": "", "industry": "",
              "segment": "Equity", "status": "Active"}
    try:
        r = requests.get(url, params=params, headers=HEADERS, timeout=30)
        data = r.json()
        rows = data if isinstance(data, list) else data.get("Table", [])
        master = {}
        for row in rows:
            code = str(row.get("SCRIP_CD") or row.get("scripCode") or "")
            name = row.get("LONG_NM") or row.get("companyName") or ""
            if code and name:
                master[normalise(name)] = code
        log.info(f"  Scrip master loaded: {len(master):,} companies")
        return master
    except Exception as e:
        log.error(f"Could not build scrip master: {e}")
        return {}


def normalise(name: str) -> str:
    """Lowercase, strip common suffixes, collapse whitespace for matching."""
    name = name.lower().strip()
    for suffix in [
        r"\blimited\b", r"\bltd\.?\b", r"\bltd\b",
        r"\bpvt\.?\b", r"\bprivate\b", r"\bco\.?\b",
        r"\bcompany\b", r"\bindustries\b", r"\bindustry\b",
        r"\benterprises\b", r"\bcorporation\b", r"\bcorp\.?\b",
    ]:
        name = re.sub(suffix, "", name)
    name = re.sub(r"[^\w\s]", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def find_scrip_code(company_name: str, master: dict) -> tuple:
    """
    Return (scrip_code, matched_name) for a company name.
    Uses exact normalised match, then fuzzy, then substring.
    Returns ('', '') if nothing found.
    """
    key = normalise(company_name)

    # 1. Exact match
    if key in master:
        return master[key], key

    # 2. Fuzzy match
    candidates = get_close_matches(key, master.keys(), n=1, cutoff=0.75)
    if candidates:
        best = candidates[0]
        log.info(f"  Fuzzy matched '{company_name}' -> '{best}'")
        return master[best], best

    # 3. Substring match
    for mkey, code in master.items():
        if key in mkey or mkey in key:
            log.info(f"  Substring matched '{company_name}' -> '{mkey}'")
            return code, mkey

    return "", ""

# ─────────────────────────── ANNOUNCEMENT FETCHER ────────────────────────────

def fetch_all_announcements(session: requests.Session, scrip_code: str) -> list:
    """Fetch ALL corporate announcements for a scrip. Handles pagination."""
    url = "https://api.bseindia.com/BseIndiaAPI/api/AnnSubCategoryGetData/w"
    all_rows = []
    page = 1

    while True:
        params = {
            "pageno":      page,
            "strCat":      "-1",
            "strPrevDate": "20140101",
            "strScrip":    scrip_code,
            "strSearch":   "P",
            "strToDate":   "20261231",
            "strType":     "C",
            "subcategory": "-1",
        }
        try:
            r = session.get(url, params=params, headers=HEADERS,
                            timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                log.warning(f"  Announcements API {r.status_code}")
                break
            data = r.json()
        except Exception as e:
            log.warning(f"  Fetch error: {e}")
            break

        rows  = data.get("Table",  [])
        meta  = data.get("Table1", [{}])
        total = int(meta[0].get("ROWCNT", 0)) if meta else 0

        all_rows.extend(rows)
        if not rows or len(all_rows) >= total:
            break

        page += 1
        time.sleep(DELAY_BETWEEN_REQUESTS)

    return all_rows

# ─────────────────────────── YEAR EXTRACTION ─────────────────────────────────

def extract_fy(subject: str, news_date: str) -> str:
    """
    Determine the financial year label (e.g. '2021-22') from subject or date.
    """
    # "2021-22" or "2021-2022"
    m = re.search(r"(20\d{2})[.\u2013\-/](20)?(\d{2})\b", subject)
    if m:
        return f"{m.group(1)}-{m.group(3)}"

    # "FY2022" or "FY 2022"
    m = re.search(r"FY\s*(\d{4})", subject, re.IGNORECASE)
    if m:
        end = int(m.group(1))
        return f"{end - 1}-{str(end)[2:]}"

    # Standalone 4-digit year
    m = re.search(r"\b(20\d{2})\b", subject)
    if m:
        end = int(m.group(1))
        return f"{end - 1}-{str(end)[2:]}"

    # Fall back to news date
    if news_date:
        try:
            dt = datetime.strptime(news_date[:10], "%Y-%m-%d")
            end = dt.year if dt.month >= 7 else dt.year - 1
            return f"{end - 1}-{str(end)[2:]}"
        except Exception:
            pass

    return ""


def is_annual_report(subject: str) -> bool:
    s = subject.lower()
    return any(kw in s for kw in ["annual report", "annual-report", "annualreport"])

# ─────────────────────────── PDF DOWNLOADER ──────────────────────────────────

def download_pdf(session: requests.Session, attachment: str, save_path: Path) -> bool:
    url = f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{attachment}"
    try:
        r = session.get(url, headers=HEADERS, timeout=60, stream=True)
        if r.status_code != 200:
            log.warning(f"  HTTP {r.status_code}: {url}")
            return False
        save_path.parent.mkdir(parents=True, exist_ok=True)
        with open(save_path, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)
        size_kb = save_path.stat().st_size / 1024
        if size_kb < 20:
            save_path.unlink()
            log.warning(f"  File too small ({size_kb:.1f} KB) — likely error page")
            return False
        log.info(f"  OK  {save_path.name}  ({size_kb:.0f} KB)")
        return True
    except Exception as e:
        log.warning(f"  Download error: {e}")
        return False

# ─────────────────────────── PER-COMPANY LOGIC ───────────────────────────────

def clean_name(name: str) -> str:
    name = name.strip().rstrip(".")
    name = re.sub(r"[^\w\s-]", "", name)
    name = re.sub(r"\s+", "_", name)
    return re.sub(r"_+", "_", name).strip("_")


def scrape_company(session, company_name, scrip_code, log_rows, done_pairs):
    file_prefix  = clean_name(company_name)
    downloaded   = set()

    announcements = fetch_all_announcements(session, scrip_code)
    log.info(f"  {len(announcements)} total announcements")

    ar_filings = [
        a for a in announcements
        if is_annual_report(a.get("NEWSSUB", "") or a.get("HEADLINE", ""))
    ]
    log.info(f"  {len(ar_filings)} annual report filings")

    for ann in ar_filings:
        subject    = ann.get("NEWSSUB", "") or ann.get("HEADLINE", "") or ""
        attachment = ann.get("ATTACHMENTNAME", "")
        news_date  = ann.get("NEWS_DT", "") or ann.get("DissemDt", "") or ""

        if not attachment or not attachment.lower().endswith(".pdf"):
            continue

        fy = extract_fy(subject, news_date)
        if not fy or fy not in YEARS or fy in downloaded:
            continue
        if (company_name, fy) in done_pairs:
            log.info(f"  skip {fy} (already done)")
            downloaded.add(fy)
            continue

        filename  = f"{file_prefix}_{fy}.pdf"
        save_path = Path(OUTPUT_DIR) / filename

        if save_path.exists():
            log.info(f"  skip {filename} (file exists)")
            downloaded.add(fy)
            log_rows.append(dict(company=company_name, year=fy,
                                 status="SKIPPED", reason="Already exists",
                                 filename=filename))
            continue

        log.info(f"  -> FY {fy}: {subject[:70]}")
        time.sleep(DELAY_BETWEEN_REQUESTS)
        ok = download_pdf(session, attachment, save_path)

        log_rows.append(dict(
            company=company_name, year=fy,
            status="SUCCESS" if ok else "FAILED",
            reason="" if ok else "Download failed",
            filename=filename if ok else "",
        ))
        if ok:
            downloaded.add(fy)

    for yr in YEARS:
        if yr not in downloaded and not any(
            r["company"] == company_name and r["year"] == yr for r in log_rows
        ):
            log_rows.append(dict(company=company_name, year=yr,
                                 status="NOT_FOUND", reason="No filing found",
                                 filename=""))

# ─────────────────────────── MAIN ────────────────────────────────────────────

def write_log(rows):
    if not rows:
        return
    with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["company","year","status","reason","filename"])
        w.writeheader()
        w.writerows(rows)


def main():
    df = pd.read_excel(EXCEL_FILE, header=None, names=["Company", "Industry"])
    companies = df["Company"].dropna().str.strip().tolist()
    log.info(f"Loaded {len(companies)} companies")
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    # KEY FIX: build scrip master ONCE from BSE's full list
    master = build_scrip_master()
    if not master:
        log.error("Scrip master is empty. Install bsedata:  pip install bsedata")
        return

    # Resume support
    log_rows   = []
    done_pairs = set()
    if Path(LOG_FILE).exists():
        with open(LOG_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row["status"] in ("SUCCESS", "SKIPPED"):
                    done_pairs.add((row["company"], row["year"]))
        log.info(f"Resuming — {len(done_pairs)} pairs already complete")

    session = requests.Session()
    session.headers.update(HEADERS)
    try:
        session.get("https://www.bseindia.com/", timeout=15)
        time.sleep(1)
    except Exception:
        pass

    no_code = []

    for idx, company in enumerate(companies, 1):
        log.info(f"\n[{idx}/{len(companies)}] {company}")

        scrip_code, _ = find_scrip_code(company, master)

        if not scrip_code:
            log.warning("  No scrip code found — skipping")
            no_code.append(company)
            for yr in YEARS:
                log_rows.append(dict(company=company, year=yr,
                                     status="FAILED",
                                     reason="Scrip code not found",
                                     filename=""))
            write_log(log_rows)
            continue

        if all((company, yr) in done_pairs for yr in YEARS):
            log.info("  All years done — skipping")
            continue

        scrape_company(session, company, scrip_code, log_rows, done_pairs)
        write_log(log_rows)
        time.sleep(DELAY_BETWEEN_COMPANIES)

    session.close()

    from collections import Counter
    counts = Counter(r["status"] for r in log_rows)
    log.info("\n" + "="*50)
    log.info(f"SUCCESS   : {counts.get('SUCCESS',   0)}")
    log.info(f"SKIPPED   : {counts.get('SKIPPED',   0)}")
    log.info(f"NOT_FOUND : {counts.get('NOT_FOUND', 0)}")
    log.info(f"FAILED    : {counts.get('FAILED',    0)}")
    log.info(f"Log: {LOG_FILE}")

    if no_code:
        log.info(f"\nNo scrip code found for {len(no_code)} companies:")
        for c in no_code:
            log.info(f"  - {c}")
        log.info("Look these up manually on bseindia.com")


if __name__ == "__main__":
    main()