"""
BSE India Annual Report Scraper
================================
Downloads annual reports (2015/16 - 2024/25) for all Nifty 500 companies
from bseindia.com

Requirements:
    pip install requests openpyxl pandas

Usage:
    python bse_annual_report_scraper.py

Output:
    ./annual_reports/<CompanyName>_<Year>.pdf
    ./bse_scraper_log.csv  (log of all successes/failures)
"""

import requests
import pandas as pd
import os
import re
import time
import csv
import json
import logging
from datetime import datetime
from pathlib import Path

# ─────────────────────────── CONFIGURATION ───────────────────────────────────

EXCEL_FILE     = "ind_nifty500list.xlsx"   # Path to your Nifty 500 Excel file
OUTPUT_DIR     = "./annual_reports"         # Where PDFs will be saved
LOG_FILE       = "./bse_scraper_log.csv"    # CSV log of results
YEARS          = [                          # Financial years to download
    "2015-16", "2016-17", "2017-18", "2018-19", "2019-20",
    "2020-21", "2021-22", "2022-23", "2023-24", "2024-25",
]

DELAY_BETWEEN_COMPANIES = 2   # seconds between companies (be polite to BSE)
DELAY_BETWEEN_REQUESTS  = 0.5 # seconds between API calls for one company
REQUEST_TIMEOUT         = 20  # seconds

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer":         "https://www.bseindia.com/",
    "Accept":          "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Origin":          "https://www.bseindia.com",
}

# ─────────────────────────── LOGGING SETUP ───────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("bse_scraper.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ─────────────────────────── HELPERS ─────────────────────────────────────────

def clean_company_name(name: str) -> str:
    """Convert company name to safe filename format."""
    name = name.strip()
    # Remove trailing punctuation like '.'
    name = name.rstrip(".")
    # Replace spaces and special chars with underscores
    name = re.sub(r"[^\w\s-]", "", name)
    name = re.sub(r"[\s]+", "_", name)
    name = re.sub(r"_+", "_", name)
    return name.strip("_")


def year_to_date_range(year: str):
    """
    '2015-16' → from_date='20150401', to_date='20160930'
    We use a wide window so we catch reports filed any time during/after the FY.
    """
    start_yr = int(year.split("-")[0])
    # Annual reports are usually filed 6-18 months after FY end
    from_date = f"{start_yr}0101"
    to_date   = f"{start_yr + 2}1231"
    return from_date, to_date


def get_scrip_code(session: requests.Session, company_name: str) -> str | None:
    """Look up BSE scrip code by company name using BSE's search API."""
    url = "https://api.bseindia.com/BseIndiaAPI/api/fetchcomp/w"
    try:
        r = session.get(url, params={"scripcode": company_name}, headers=HEADERS,
                        timeout=REQUEST_TIMEOUT)
        if r.status_code != 200:
            return None
        data = r.json()
        # Response: list of {SCRIP_CD, FULL_NM, ...}
        if isinstance(data, list) and data:
            return str(data[0].get("SCRIP_CD", ""))
        return None
    except Exception as e:
        log.warning(f"  Scrip lookup failed for '{company_name}': {e}")
        return None


def search_scrip_code(session: requests.Session, company_name: str) -> str | None:
    """Alternative: use BSE's typeahead/search endpoint."""
    url = "https://api.bseindia.com/BseIndiaAPI/api/ComHeadernew/w"
    try:
        r = session.get(url, params={"code": company_name}, headers=HEADERS,
                        timeout=REQUEST_TIMEOUT)
        if r.status_code != 200:
            return None
        data = r.json()
        # Returns list of matches
        if isinstance(data, list) and data:
            return str(data[0].get("SCRIP_CD", ""))
        return None
    except Exception as e:
        log.warning(f"  Search scrip lookup failed for '{company_name}': {e}")
        return None


def get_scrip_code_robust(session: requests.Session, company_name: str) -> str | None:
    """Try multiple BSE endpoints to find the scrip code."""
    endpoints = [
        # Endpoint 1: direct search
        ("https://api.bseindia.com/BseIndiaAPI/api/fetchcomp/w",
         {"scripcode": company_name}),
        # Endpoint 2: typeahead search
        ("https://api.bseindia.com/BseIndiaAPI/api/ComHeadernew/w",
         {"code": company_name}),
        # Endpoint 3: list with search
        ("https://api.bseindia.com/BseIndiaAPI/api/ListofScripData/w",
         {"Group": "", "Scripcode": "", "industry": "",
          "segment": "Equity", "status": "Active", "companyname": company_name}),
    ]
    for url, params in endpoints:
        try:
            r = session.get(url, params=params, headers=HEADERS,
                            timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                continue
            data = r.json()
            # Handle list response
            if isinstance(data, list) and data:
                code = data[0].get("SCRIP_CD") or data[0].get("scripCode")
                if code:
                    return str(code)
            # Handle dict with Table key
            if isinstance(data, dict):
                table = data.get("Table") or data.get("data") or []
                if table:
                    code = table[0].get("SCRIP_CD") or table[0].get("scripCode")
                    if code:
                        return str(code)
            time.sleep(0.3)
        except Exception as e:
            log.debug(f"  Endpoint {url} failed: {e}")
    return None


def fetch_announcements(session: requests.Session, scrip_code: str,
                         from_date: str, to_date: str) -> list[dict]:
    """
    Fetch all corporate announcements for a scrip within a date range.
    Paginates automatically until all pages are fetched.
    """
    url = "https://api.bseindia.com/BseIndiaAPI/api/AnnSubCategoryGetData/w"
    all_rows = []
    page = 1

    while True:
        params = {
            "pageno":      page,
            "strCat":      "-1",       # all categories
            "strPrevDate": from_date,
            "strScrip":    scrip_code,
            "strSearch":   "P",
            "strToDate":   to_date,
            "strType":     "C",
            "subcategory": "-1",
        }
        try:
            r = session.get(url, params=params, headers=HEADERS,
                            timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                log.warning(f"  Announcements API returned {r.status_code}")
                break
            data = r.json()
        except Exception as e:
            log.warning(f"  Announcements fetch error: {e}")
            break

        rows  = data.get("Table",  [])
        meta  = data.get("Table1", [{}])
        total = int(meta[0].get("ROWCNT", 0)) if meta else 0

        all_rows.extend(rows)

        if len(all_rows) >= total or not rows:
            break
        page += 1
        time.sleep(DELAY_BETWEEN_REQUESTS)

    return all_rows


def is_annual_report(subject: str) -> bool:
    """Return True if the announcement subject looks like an annual report."""
    subject_lower = subject.lower()
    keywords = ["annual report", "annual-report", "annualreport"]
    return any(kw in subject_lower for kw in keywords)


def extract_fy_from_subject(subject: str, default_from_date: str) -> str | None:
    """
    Try to extract the financial year label from the announcement subject.
    e.g. 'Annual Report 2021-22' → '2021-22'
         'Annual Report for FY 2022' → '2021-22'
    Falls back to deriving year from from_date.
    """
    # Pattern 1: "2021-22" or "2021-2022"
    m = re.search(r"(20\d{2})[-–/](20)?(\d{2})", subject)
    if m:
        start = m.group(1)
        end   = m.group(3)
        return f"{start}-{end}"

    # Pattern 2: "FY2022" or "FY 2022"
    m = re.search(r"FY\s*(\d{4})", subject, re.IGNORECASE)
    if m:
        end_yr   = int(m.group(1))
        start_yr = end_yr - 1
        return f"{start_yr}-{str(end_yr)[2:]}"

    # Pattern 3: standalone year "2022"
    m = re.search(r"\b(20\d{2})\b", subject)
    if m:
        end_yr   = int(m.group(1))
        start_yr = end_yr - 1
        return f"{start_yr}-{str(end_yr)[2:]}"

    return None


def download_pdf(session: requests.Session, attachment: str,
                 save_path: Path) -> bool:
    """Download a PDF from BSE's corpfiling store. Returns True on success."""
    pdf_url = (
        f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{attachment}"
    )
    try:
        r = session.get(pdf_url, headers=HEADERS, timeout=60, stream=True)
        if r.status_code != 200:
            log.warning(f"  PDF download failed (HTTP {r.status_code}): {pdf_url}")
            return False
        # Check it's actually a PDF
        content_type = r.headers.get("Content-Type", "")
        if "pdf" not in content_type.lower() and not attachment.endswith(".pdf"):
            log.warning(f"  Not a PDF (Content-Type={content_type}): {pdf_url}")
            return False

        save_path.parent.mkdir(parents=True, exist_ok=True)
        with open(save_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)

        size_kb = save_path.stat().st_size / 1024
        if size_kb < 10:
            # Suspiciously small — likely an error HTML page
            save_path.unlink()
            log.warning(f"  Downloaded file too small ({size_kb:.1f} KB), discarding.")
            return False

        log.info(f"  ✓ Saved ({size_kb:.0f} KB): {save_path.name}")
        return True

    except Exception as e:
        log.warning(f"  PDF download exception: {e}")
        return False


# ─────────────────────────── MAIN SCRAPER ────────────────────────────────────

def scrape_company(session: requests.Session, company_name: str,
                   clean_name: str, log_rows: list) -> None:
    """Scrape all annual reports for a single company across all target years."""
    log.info(f"\n{'─'*60}")
    log.info(f"Company: {company_name}")

    # ── Step 1: Get scrip code ────────────────────────────────────────────────
    scrip_code = get_scrip_code_robust(session, company_name)
    if not scrip_code:
        log.warning(f"  ✗ Could not find scrip code — skipping.")
        for yr in YEARS:
            log_rows.append({
                "company":    company_name,
                "year":       yr,
                "status":     "FAILED",
                "reason":     "Scrip code not found",
                "filename":   "",
                "url":        "",
            })
        return

    log.info(f"  Scrip Code: {scrip_code}")

    # ── Step 2: Fetch announcements for full period ───────────────────────────
    # We fetch a wide range once, then match by year in subject/date
    from_date = f"{int(YEARS[0].split('-')[0]) - 1}0101"   # 1 year before first FY
    to_date   = f"{int(YEARS[-1].split('-')[0]) + 2}1231"  # 1 year after last FY

    time.sleep(DELAY_BETWEEN_REQUESTS)
    announcements = fetch_announcements(session, scrip_code, from_date, to_date)
    log.info(f"  Found {len(announcements)} total announcements")

    # ── Step 3: Filter to annual report announcements ─────────────────────────
    ar_filings = []
    for ann in announcements:
        subject = ann.get("NEWSSUB", "") or ann.get("HEADLINE", "") or ""
        if is_annual_report(subject):
            ar_filings.append(ann)

    log.info(f"  Found {len(ar_filings)} annual report filings")

    if not ar_filings:
        log.warning("  No annual report filings found in announcements.")
        for yr in YEARS:
            log_rows.append({
                "company":  company_name,
                "year":     yr,
                "status":   "NOT_FOUND",
                "reason":   "No annual report announcements found",
                "filename": "",
                "url":      "",
            })
        return

    # ── Step 4: Match each filing to a target year and download ──────────────
    downloaded_years = set()

    for ann in ar_filings:
        subject    = ann.get("NEWSSUB", "") or ann.get("HEADLINE", "") or ""
        attachment = ann.get("ATTACHMENTNAME", "")
        news_date  = ann.get("NEWS_DT", "") or ann.get("DissemDt", "") or ""

        if not attachment or not attachment.endswith(".pdf"):
            continue

        # Try to figure out which FY this report is for
        fy = extract_fy_from_subject(subject, "")

        # If we can't extract FY from subject, infer from filing date
        if not fy and news_date:
            try:
                dt = datetime.strptime(news_date[:10], "%Y-%m-%d")
                # Annual reports are filed Apr–Dec; FY ended in March of that year
                if dt.month >= 4:
                    end_yr   = dt.year
                else:
                    end_yr   = dt.year - 1
                start_yr = end_yr - 1
                fy = f"{start_yr}-{str(end_yr)[2:]}"
            except Exception:
                pass

        if not fy or fy not in YEARS:
            continue

        if fy in downloaded_years:
            continue  # Already got this year

        safe_filename = f"{clean_name}_{fy}.pdf"
        save_path     = Path(OUTPUT_DIR) / safe_filename
        pdf_url       = (
            f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{attachment}"
        )

        if save_path.exists():
            log.info(f"  ⤵  Already exists, skipping: {safe_filename}")
            downloaded_years.add(fy)
            log_rows.append({
                "company":  company_name,
                "year":     fy,
                "status":   "SKIPPED",
                "reason":   "File already exists",
                "filename": safe_filename,
                "url":      pdf_url,
            })
            continue

        log.info(f"  Downloading FY {fy}: {subject[:70]}")
        time.sleep(DELAY_BETWEEN_REQUESTS)
        success = download_pdf(session, attachment, save_path)

        log_rows.append({
            "company":  company_name,
            "year":     fy,
            "status":   "SUCCESS" if success else "FAILED",
            "reason":   "" if success else "PDF download failed",
            "filename": safe_filename if success else "",
            "url":      pdf_url,
        })

        if success:
            downloaded_years.add(fy)

    # ── Step 5: Log any target years that weren't found ───────────────────────
    for yr in YEARS:
        if yr not in downloaded_years:
            already_logged = any(
                r["company"] == company_name and r["year"] == yr
                for r in log_rows
            )
            if not already_logged:
                log_rows.append({
                    "company":  company_name,
                    "year":     yr,
                    "status":   "NOT_FOUND",
                    "reason":   "No filing found for this year",
                    "filename": "",
                    "url":      "",
                })


def main():
    # ── Load company list ─────────────────────────────────────────────────────
    df = pd.read_excel(EXCEL_FILE, header=None, names=["Company", "Industry"])
    companies = df["Company"].dropna().str.strip().tolist()
    log.info(f"Loaded {len(companies)} companies from {EXCEL_FILE}")
    log.info(f"Target years: {YEARS[0]} → {YEARS[-1]}")
    log.info(f"Output directory: {OUTPUT_DIR}\n")

    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    log_rows = []

    # ── Resume support: load existing log to skip done companies ─────────────
    done_pairs = set()
    if Path(LOG_FILE).exists():
        with open(LOG_FILE, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("status") in ("SUCCESS", "SKIPPED"):
                    done_pairs.add((row["company"], row["year"]))
        log.info(f"Resuming: {len(done_pairs)} already-completed (company, year) pairs found.")

    # ── HTTP session ──────────────────────────────────────────────────────────
    session = requests.Session()
    session.headers.update(HEADERS)

    # Warm up the session — BSE sometimes needs a browser-like first request
    try:
        session.get("https://www.bseindia.com/", timeout=15)
        time.sleep(1)
    except Exception:
        pass

    # ── Process each company ──────────────────────────────────────────────────
    for idx, company in enumerate(companies, 1):
        clean_name = clean_company_name(company)

        # Skip if all years already done for this company
        all_done = all((company, yr) in done_pairs for yr in YEARS)
        if all_done:
            log.info(f"[{idx}/{len(companies)}] Skipping (all years done): {company}")
            continue

        log.info(f"\n[{idx}/{len(companies)}]")
        scrape_company(session, company, clean_name, log_rows)

        # Write log incrementally (so progress is saved even if script crashes)
        _write_log(log_rows)

        time.sleep(DELAY_BETWEEN_COMPANIES)

    session.close()

    log.info("\n" + "="*60)
    log.info("SCRAPING COMPLETE")
    _print_summary(log_rows)
    log.info(f"Full log saved to: {LOG_FILE}")


def _write_log(log_rows: list) -> None:
    """Write/overwrite the CSV log file."""
    if not log_rows:
        return
    with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["company", "year", "status", "reason", "filename", "url"]
        )
        writer.writeheader()
        writer.writerows(log_rows)


def _print_summary(log_rows: list) -> None:
    from collections import Counter
    counts = Counter(r["status"] for r in log_rows)
    log.info(f"  SUCCESS   : {counts.get('SUCCESS',   0)}")
    log.info(f"  SKIPPED   : {counts.get('SKIPPED',   0)}")
    log.info(f"  NOT_FOUND : {counts.get('NOT_FOUND', 0)}")
    log.info(f"  FAILED    : {counts.get('FAILED',    0)}")


if __name__ == "__main__":
    main()