"""
═══════════════════════════════════════════════════════════════════════════════
  TAVE EVENT STUDY — CODE 1 of 2
  EVENT COLLECTION, CLASSIFICATION & DATA ASSEMBLY
═══════════════════════════════════════════════════════════════════════════════

  Purpose:
    1. Collect tokenization announcements for 50 firms across 6 markets
       (US / SGX / XETRA / LSE / SIX / HKEX)
    2. Auto-classify each event: type, asset class, jurisdiction,
       audit disclosure, % balance sheet tokenized
    3. Pull financials from SimFin (paid) + prices from yfinance
    4. Export ONE master Excel file to Google Drive

  Output: tave_master_database.xlsx with sheets:
    - Events_Classified   (all events + classification)
    - Firm_Financials     (SimFin fundamentals)
    - Price_Coverage      (data availability check)
    - Manual_Review       (low-confidence classifications flagged)

  Author: Felix Diego Langer — TAVE Research, GlobalNxt DBA 2026
  Run on: Google Colab (see setup document)
═══════════════════════════════════════════════════════════════════════════════
"""

# ─────────────────────────────────────────────────────────────────────────────
# INSTALL (run in Colab first cell):
#   !pip install requests pandas numpy tqdm beautifulsoup4 lxml openpyxl yfinance simfin
# ─────────────────────────────────────────────────────────────────────────────

import requests
import pandas as pd
import numpy as np
import time
import re
import os
from datetime import datetime
from tqdm import tqdm
import warnings
warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════

START_DATE = "2021-01-01"
END_DATE   = "2025-12-31"

# Your SimFin API key — set in Colab via:  os.environ["SIMFIN_KEY"] = "your_key"
def _get_secret(name, default=""):
    """
    Read a secret from Colab's secret manager (the key icon in the left sidebar)
    if available, else fall back to an environment variable, else the default.
    This keeps your SimFin key and email OUT of the committed code.
    """
    # 1. Colab secrets (userdata)
    try:
        from google.colab import userdata
        val = userdata.get(name)
        if val:
            return val
    except Exception:
        pass
    # 2. Environment variable
    return os.environ.get(name, default)


# SimFin key — set as a Colab secret named SIMFIN_KEY (recommended), or env var
SIMFIN_API_KEY = _get_secret("SIMFIN_KEY", "PUT_YOUR_KEY_HERE")

# Google Drive output path (mounted in Colab)
OUTPUT_DIR  = "/content/drive/MyDrive/TAVE_Research"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "tave_master_database.xlsx")

# SEC requires a declared User-Agent with contact info (their fair-access policy).
# Set a Colab secret named SEC_EMAIL to keep your address out of committed code.
# Falls back to a generic placeholder if not set (works, but set your real email).
SEC_EMAIL = _get_secret("SEC_EMAIL", "")
SEC_USER_AGENT = f"TAVE Research {SEC_EMAIL}" if SEC_EMAIL else "TAVE Academic Research contact-via-github"



def ensure_drive_and_dirs():
    """
    Make running Code 1 'just work' like the CVM setup.
    - If on Colab and Drive isn't mounted yet, mount it.
    - Create the TAVE_Research output folder if it doesn't exist.
    - Warn (do not crash) if the SimFin key or SEC email still hold placeholders.
    Safe to call anywhere; silently no-ops off Colab.
    """
    # 1. Mount Drive if we're on Colab and it isn't mounted
    if not os.path.isdir("/content/drive/MyDrive"):
        try:
            from google.colab import drive   # only exists on Colab
            print("  Mounting Google Drive...")
            drive.mount("/content/drive")
        except ImportError:
            # Not on Colab — fall back to a local folder so the script still runs
            global OUTPUT_DIR, OUTPUT_FILE
            if not os.path.isdir("/content/drive/MyDrive"):
                local = os.path.join(os.getcwd(), "TAVE_Research")
                OUTPUT_DIR = local
                OUTPUT_FILE = os.path.join(OUTPUT_DIR, "tave_master_database.xlsx")
                print(f"  Not on Colab — writing output locally to: {OUTPUT_DIR}")
        except Exception as e:
            print(f"  Drive mount issue ({e}); will attempt to write to {OUTPUT_DIR} anyway.")

    # 2. Create the output folder
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"  Output folder ready: {OUTPUT_DIR}")

    # 3. Config sanity warnings (non-fatal)
    if SIMFIN_API_KEY in ("", "PUT_YOUR_KEY_HERE", None):
        print("  ⚠ SIMFIN_KEY not found. Add it as a Colab secret (key icon, left sidebar)")
        print("    named SIMFIN_KEY, or run: import os; os.environ['SIMFIN_KEY']='your_key'")
        print("    (Collection still works; only SimFin financials will be skipped.)")
    if not SEC_EMAIL:
        print("  ⚠ SEC_EMAIL not set. EDGAR works without it, but SEC's fair-access policy")
        print("    prefers a real contact. Add a Colab secret named SEC_EMAIL with your email.")


TOKENIZATION_KEYWORDS = [
    "tokeniz", "tokenis", "digital asset", "blockchain", "distributed ledger",
    "smart contract", "real world asset", "real-world asset", "on-chain",
    "on chain", "digital bond", "digital securities", "tokenized fund",
    "tokenised fund", "DLT", "asset tokeniz",
]

# ═══════════════════════════════════════════════════════════════════════════════
# FIRM UNIVERSE — 50 FIRMS
# ═══════════════════════════════════════════════════════════════════════════════

FIRMS = {
    # ── US (10) — SEC EDGAR ──────────────────────────────────────────────────
    "US": [
        {"name": "JPMorgan Chase",      "ticker": "JPM",  "yf": "JPM",  "cik": "0000019617", "jurisdiction": "US"},
        {"name": "BlackRock",           "ticker": "BLK",  "yf": "BLK",  "cik": "0001364742", "jurisdiction": "US"},
        {"name": "Franklin Resources",  "ticker": "BEN",  "yf": "BEN",  "cik": "0000038777", "jurisdiction": "US"},
        {"name": "Goldman Sachs",       "ticker": "GS",   "yf": "GS",   "cik": "0000886982", "jurisdiction": "US"},
        {"name": "Citigroup",           "ticker": "C",    "yf": "C",    "cik": "0000831001", "jurisdiction": "US"},
        {"name": "Bank of America",     "ticker": "BAC",  "yf": "BAC",  "cik": "0000070858", "jurisdiction": "US"},
        {"name": "Morgan Stanley",      "ticker": "MS",   "yf": "MS",   "cik": "0000895421", "jurisdiction": "US"},
        {"name": "Nasdaq Inc",          "ticker": "NDAQ", "yf": "NDAQ", "cik": "0001120193", "jurisdiction": "US"},
        {"name": "State Street",        "ticker": "STT",  "yf": "STT",  "cik": "0000093751", "jurisdiction": "US"},
        {"name": "Visa Inc",            "ticker": "V",    "yf": "V",    "cik": "0001403161", "jurisdiction": "US"},
    ],
    # ── SINGAPORE (8) — SGX API ──────────────────────────────────────────────
    "SGX": [
        {"name": "DBS Group Holdings",  "ticker": "D05",  "yf": "D05.SI",  "stock_id": "D05",  "jurisdiction": "Singapore"},
        {"name": "Singapore Exchange",  "ticker": "S68",  "yf": "S68.SI",  "stock_id": "S68",  "jurisdiction": "Singapore"},
        {"name": "OCBC Bank",           "ticker": "O39",  "yf": "O39.SI",  "stock_id": "O39",  "jurisdiction": "Singapore"},
        {"name": "United Overseas Bank","ticker": "U11",  "yf": "U11.SI",  "stock_id": "U11",  "jurisdiction": "Singapore"},
        {"name": "CapitaLand Invest",   "ticker": "9CI",  "yf": "9CI.SI",  "stock_id": "9CI",  "jurisdiction": "Singapore"},
        {"name": "Keppel Ltd",          "ticker": "BN4",  "yf": "BN4.SI",  "stock_id": "BN4",  "jurisdiction": "Singapore"},
        {"name": "Sea Limited",         "ticker": "SE",   "yf": "SE",      "stock_id": None,   "jurisdiction": "Singapore", "us_listed": True, "cik": "0001703399"},
        {"name": "Mapletree Pan Asia",  "ticker": "N2IU", "yf": "N2IU.SI", "stock_id": "N2IU", "jurisdiction": "Singapore"},
    ],
    # ── GERMANY (5) — EQS/DGAP ───────────────────────────────────────────────
    "XETRA": [
        {"name": "Siemens AG",          "ticker": "SIE",  "yf": "SIE.DE", "isin": "DE0007236101", "jurisdiction": "EU"},
        {"name": "Deutsche Bank",       "ticker": "DBK",  "yf": "DBK.DE", "isin": "DE0005140008", "jurisdiction": "EU"},
        {"name": "Commerzbank",         "ticker": "CBK",  "yf": "CBK.DE", "isin": "DE000CBK1001", "jurisdiction": "EU"},
        {"name": "DWS Group",           "ticker": "DWS",  "yf": "DWS.DE", "isin": "DE000DWS1007", "jurisdiction": "EU"},
        {"name": "Deutsche Boerse",     "ticker": "DB1",  "yf": "DB1.DE", "isin": "DE0005810055", "jurisdiction": "EU"},
    ],
    # ── UK (6) — RNS/LSE ─────────────────────────────────────────────────────
    "LSE": [
        {"name": "HSBC Holdings",       "ticker": "HSBA", "yf": "HSBA.L", "isin": "GB0005405286", "jurisdiction": "UK"},
        {"name": "Barclays",            "ticker": "BARC", "yf": "BARC.L", "isin": "GB0031348658", "jurisdiction": "UK"},
        {"name": "Standard Chartered",  "ticker": "STAN", "yf": "STAN.L", "isin": "GB0004082847", "jurisdiction": "UK"},
        {"name": "LSE Group",           "ticker": "LSEG", "yf": "LSEG.L", "isin": "GB00B0SWJX34", "jurisdiction": "UK"},
        {"name": "Schroders",           "ticker": "SDR",  "yf": "SDR.L",  "isin": "GB0002405495", "jurisdiction": "UK"},
        {"name": "Lloyds Banking",      "ticker": "LLOY", "yf": "LLOY.L", "isin": "GB0008706128", "jurisdiction": "UK"},
    ],
    # ── SWITZERLAND (5) — SIX ────────────────────────────────────────────────
    "SIX": [
        {"name": "UBS Group",           "ticker": "UBSG", "yf": "UBSG.SW", "isin": "CH0244767585", "jurisdiction": "Switzerland"},
        {"name": "Julius Baer",         "ticker": "BAER", "yf": "BAER.SW", "isin": "CH0102484968", "jurisdiction": "Switzerland"},
        {"name": "SIX Swiss (Pfd proxy)","ticker":"SIXN", "yf": None,      "isin": "CH0362432707", "jurisdiction": "Switzerland"},
        {"name": "Zurich Insurance",    "ticker": "ZURN", "yf": "ZURN.SW", "isin": "CH0011075394", "jurisdiction": "Switzerland"},
        {"name": "Partners Group",      "ticker": "PGHN", "yf": "PGHN.SW", "isin": "CH0024608827", "jurisdiction": "Switzerland"},
    ],
    # ── HONG KONG (8) — HKEX ─────────────────────────────────────────────────
    "HKEX": [
        {"name": "HSBC Holdings HK",    "ticker": "0005", "yf": "0005.HK", "hkex_id": "00005", "jurisdiction": "Hong Kong"},
        {"name": "HK Exchanges",        "ticker": "0388", "yf": "0388.HK", "hkex_id": "00388", "jurisdiction": "Hong Kong"},
        {"name": "Bank of China HK",    "ticker": "2388", "yf": "2388.HK", "hkex_id": "02388", "jurisdiction": "Hong Kong"},
        {"name": "Hang Seng Bank",      "ticker": "0011", "yf": "0011.HK", "hkex_id": "00011", "jurisdiction": "Hong Kong"},
        {"name": "CITIC Securities",    "ticker": "6030", "yf": "6030.HK", "hkex_id": "06030", "jurisdiction": "Hong Kong"},
        {"name": "Bank of East Asia",   "ticker": "0023", "yf": "0023.HK", "hkex_id": "00023", "jurisdiction": "Hong Kong"},
        {"name": "Standard Chartered HK","ticker":"2888", "yf": "2888.HK", "hkex_id": "02888", "jurisdiction": "Hong Kong"},
        {"name": "AIA Group",           "ticker": "1299", "yf": "1299.HK", "hkex_id": "01299", "jurisdiction": "Hong Kong"},
    ],
    # ── JAPAN (8) — TDnet (automated JSON endpoint) ──────────────────────────
    "TSE": [
        {"name": "Nomura Holdings",     "ticker": "8604", "yf": "8604.T", "edinet": "E03814", "jurisdiction": "Japan"},
        {"name": "MUFG",                "ticker": "8306", "yf": "8306.T", "edinet": "E03606", "jurisdiction": "Japan"},
        {"name": "SBI Holdings",        "ticker": "8473", "yf": "8473.T", "edinet": "E08957", "jurisdiction": "Japan"},
        {"name": "Sumitomo Mitsui FG",  "ticker": "8316", "yf": "8316.T", "edinet": "E03665", "jurisdiction": "Japan"},
        {"name": "Mizuho FG",           "ticker": "8411", "yf": "8411.T", "edinet": "E03615", "jurisdiction": "Japan"},
        {"name": "Daiwa Securities",    "ticker": "8601", "yf": "8601.T", "edinet": "E03811", "jurisdiction": "Japan"},
        {"name": "Japan Exchange Grp",  "ticker": "8697", "yf": "8697.T", "edinet": "E03814", "jurisdiction": "Japan"},
        {"name": "Mitsubishi Corp",     "ticker": "8058", "yf": "8058.T", "edinet": "E02528", "jurisdiction": "Japan"},
    ],
}

# Jurisdiction regulatory-clarity score (for the cross-sectional regression)
# Higher = clearer/more supportive regulatory framework for tokenization
JURISDICTION_SCORE = {
    "Singapore": 5, "Switzerland": 5, "EU": 4, "UK": 4,
    "Hong Kong": 4, "Japan": 3, "US": 2,
}

# ═══════════════════════════════════════════════════════════════════════════════
# CLASSIFICATION ENGINE
# ═══════════════════════════════════════════════════════════════════════════════
#
# Auto-classifies each announcement along the dimensions the regression needs.
# Uses keyword/regex matching on the headline + available body text.
# Each classification carries a confidence score; low-confidence rows are
# flagged for manual review rather than silently guessed.
# ═══════════════════════════════════════════════════════════════════════════════

# Announcement TYPE — ordered by specificity (first match wins)
# Patterns use stem+\w* so inflected forms match (launch/launches/launched/launching)
TYPE_PATTERNS = [
    ("Live Launch", r"\b(launch\w*|go[- ]?live|now available|commercially available|fully operational|first issuance|issu\w+|complet\w+|priced|settl\w+|mint\w*|deploy\w*|live)\b"),
    ("Expansion",   r"\b(expand\w*|extend\w*|scal\w+|next phase|phase 2|phase ii|additional|broaden\w*|new market)\b"),
    ("Pilot",       r"\b(pilot\w*|trial\w*|proof[- ]of[- ]concept|poc|sandbox|test\w*|experiment\w*)\b"),
    ("New Program", r"\b(launch\w*|introduc\w+|unveil\w*|announc\w+|new (platform|service|initiative|programme|program))\b"),
]

# ASSET CLASS — what is being tokenized (stem-based)
ASSET_CLASS_PATTERNS = [
    ("Bond",         r"\b(bond\w*|note[s]?|fixed income|debt security|coupon\w*|debenture\w*|gilt\w*|treasur\w+)\b"),
    ("Fund",         r"\b(fund\w*|money market|mmf|etf|mutual fund|investment fund|unit trust)\b"),
    ("Real Estate",  r"\b(real estate|propert\w+|reit\w*|commercial property|building\w*)\b"),
    ("Receivables",  r"\b(receivabl\w+|trade finance|invoice\w*|factoring|supply chain finance)\b"),
    ("Supply Chain", r"\b(supply chain|logistics|inventor\w+|commodit\w+|trade asset\w*)\b"),
    ("Deposit",      r"\b(deposit\w*|tokeniz\w+ deposit|cash|stablecoin\w*|e-money)\b"),
    ("Equity",       r"\b(equit\w+|share[s]?|stock\w*|private equity|pe fund)\b"),
    ("Gold/Commodity",r"\b(gold|silver|commodit\w+|precious metal\w*|carbon credit\w*)\b"),
]

# AUDIT disclosure
AUDIT_PATTERNS = r"\b(audit|audited|security review|smart contract audit|penetration test|certik|openzeppelin|halborn|trail of bits|formal verification)\b"

# % balance sheet tokenized — capture explicit percentages or dollar amounts
PCT_PATTERN = r"(\d{1,3}(?:\.\d+)?)\s?%"
AMOUNT_PATTERN = r"(?:USD|EUR|SGD|GBP|CHF|HKD|JPY|\$|€|£)\s?(\d+(?:[.,]\d+)?)\s?(billion|bn|million|mn|m|b)\b"


def classify_event(headline: str, body: str = "") -> dict:
    """
    Classify a single announcement. Returns classification + confidence.
    Confidence is LOW if no clear match — those rows get flagged for manual review.
    """
    text = f"{headline} {body}".lower()
    result = {
        "announcement_type":   "Unclassified",
        "asset_class":         "Unclassified",
        "audit_disclosed":     False,
        "pct_tokenized":       None,
        "amount_disclosed":    None,
        "classification_conf": "LOW",
    }

    if not text.strip():
        return result

    matches_found = 0

    # Announcement type
    for label, pattern in TYPE_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            result["announcement_type"] = label
            matches_found += 1
            break

    # Asset class
    for label, pattern in ASSET_CLASS_PATTERNS:
        if re.search(pattern, text, re.IGNORECASE):
            result["asset_class"] = label
            matches_found += 1
            break

    # Audit
    if re.search(AUDIT_PATTERNS, text, re.IGNORECASE):
        result["audit_disclosed"] = True
        matches_found += 1

    # Percentage tokenized
    pct_matches = re.findall(PCT_PATTERN, text)
    if pct_matches:
        # Take the first plausible percentage (0-100)
        for p in pct_matches:
            val = float(p)
            if 0 < val <= 100:
                result["pct_tokenized"] = val
                break

    # Dollar amount (proxy for scope when % not disclosed)
    amt_match = re.search(AMOUNT_PATTERN, text, re.IGNORECASE)
    if amt_match:
        num = float(amt_match.group(1).replace(",", "."))
        unit = amt_match.group(2).lower()
        multiplier = 1e9 if unit in ("billion", "bn", "b") else 1e6
        result["amount_disclosed"] = num * multiplier

    # Confidence: HIGH if both type AND asset class classified, MEDIUM if one, LOW if none
    if result["announcement_type"] != "Unclassified" and result["asset_class"] != "Unclassified":
        result["classification_conf"] = "HIGH"
    elif matches_found >= 1:
        result["classification_conf"] = "MEDIUM"
    else:
        result["classification_conf"] = "LOW"

    return result


# ═══════════════════════════════════════════════════════════════════════════════
# COLLECTORS — one per market
# ═══════════════════════════════════════════════════════════════════════════════

def _date_in_range(date_str: str) -> bool:
    try:
        d = datetime.strptime(date_str[:10], "%Y-%m-%d")
        return datetime.strptime(START_DATE, "%Y-%m-%d") <= d <= datetime.strptime(END_DATE, "%Y-%m-%d")
    except (ValueError, TypeError):
        return False


def _base_row(firm, kw, date, form, accession, source, url, headline, body=""):
    """Build a standardized event row with classification applied."""
    cls = classify_event(headline, body)
    return {
        "firm_name":          firm["name"],
        "ticker":             firm["ticker"],
        "yf_ticker":          firm.get("yf", ""),
        "exchange":           firm.get("exchange", source),
        "jurisdiction":       firm.get("jurisdiction", ""),
        "jurisdiction_score": JURISDICTION_SCORE.get(firm.get("jurisdiction", ""), np.nan),
        "keyword":            kw,
        "event_date":         date,
        "form_type":          form,
        "accession_no":       accession,
        "source":             source,
        "url":                url,
        "headline":           headline,
        "announcement_type":  cls["announcement_type"],
        "asset_class":        cls["asset_class"],
        "audit_disclosed":    cls["audit_disclosed"],
        "pct_tokenized":      cls["pct_tokenized"],
        "amount_disclosed":   cls["amount_disclosed"],
        "classification_conf":cls["classification_conf"],
        "verified":           False,
        "manual_notes":       "",
    }


# ── US: SEC EDGAR ────────────────────────────────────────────────────────────
def collect_edgar(firm) -> list:
    rows = []
    endpoint = "https://efts.sec.gov/LATEST/search-index"
    headers = {"User-Agent": SEC_USER_AGENT}
    seen = set()

    for kw in TOKENIZATION_KEYWORDS:
        try:
            params = {
                "q": f'"{kw}"', "dateRange": "custom",
                "startdt": START_DATE, "enddt": END_DATE,
                "entity": firm["name"], "forms": "8-K,6-K,20-F,10-K,10-Q",
            }
            r = requests.get(endpoint, params=params, headers=headers, timeout=20)
            if r.status_code != 200:
                time.sleep(0.2); continue
            hits = r.json().get("hits", {}).get("hits", [])
            for hit in hits:
                src = hit.get("_source", {})
                acc = hit.get("_id", "").split(":")[0]
                if acc in seen:
                    continue
                seen.add(acc)
                date = src.get("file_date", "")
                if not _date_in_range(date):
                    continue
                cik_num = firm["cik"].lstrip("0")
                url = f"https://www.sec.gov/Archives/edgar/data/{cik_num}/{acc.replace('-','')}/{acc}-index.htm"
                rows.append(_base_row(
                    firm, kw, date, src.get("form_type", ""), acc,
                    "EDGAR", url, src.get("entity_name", firm["name"]),
                ))
            time.sleep(0.15)
        except Exception as e:
            print(f"    EDGAR {firm['name']}/{kw}: {e}")
    return rows


# ── SGX ──────────────────────────────────────────────────────────────────────
def collect_sgx(firm) -> list:
    rows = []
    if firm.get("us_listed"):   # e.g. Sea Ltd files with SEC
        return collect_edgar(firm)
    if not firm.get("stock_id"):
        return rows
    url = "https://api.sgx.com/securities/v1.1/announcements"
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "application/json"}
    try:
        params = {"pageStart": "0", "pageSize": "1000", "stockId": firm["stock_id"],
                  "sortBy": "announcedOn", "orderBy": "desc"}
        r = requests.get(url, params=params, headers=headers, timeout=25)
        if r.status_code == 200:
            anns = r.json().get("data", {}).get("announcements", [])
            for a in anns:
                head = a.get("headline", "")
                if not _date_in_range(str(a.get("announcedOn", ""))):
                    continue
                if not any(k.lower() in head.lower() for k in TOKENIZATION_KEYWORDS):
                    continue
                rows.append(_base_row(
                    firm, "match", str(a.get("announcedOn",""))[:10],
                    a.get("category",""), str(a.get("id","")),
                    "SGX", a.get("fileUrl",""), head,
                ))
    except Exception as e:
        print(f"    SGX {firm['name']}: {e}")
    time.sleep(0.4)
    return rows


# ── XETRA (EQS/DGAP) ─────────────────────────────────────────────────────────
def collect_eqs(firm) -> list:
    rows = []
    search_url = "https://newsfeed.eqs.com/api/news/search"
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "application/json"}
    seen = set()
    for kw in TOKENIZATION_KEYWORDS[:8]:
        try:
            params = {"query": kw, "isin": firm["isin"],
                      "dateFrom": START_DATE, "dateTo": END_DATE,
                      "language": "en", "size": 50}
            r = requests.get(search_url, params=params, headers=headers, timeout=20)
            if r.status_code == 200:
                items = r.json().get("items", []) or r.json().get("news", [])
                for it in items:
                    head = it.get("title", it.get("headline", ""))
                    date = str(it.get("publishedAt", it.get("date","")))[:10]
                    key = f"{date}_{head[:40]}"
                    if key in seen or not _date_in_range(date):
                        continue
                    seen.add(key)
                    rows.append(_base_row(firm, kw, date, "RegNews",
                                          str(it.get("id","")), "EQS",
                                          it.get("url", it.get("link","")), head))
            time.sleep(0.25)
        except Exception as e:
            print(f"    EQS {firm['name']}/{kw}: {e}")
    return rows


# ── LSE (RNS) ────────────────────────────────────────────────────────────────
def collect_rns(firm) -> list:
    rows = []
    url = "https://api.londonstockexchange.com/api/gw/lse/regulatory-news/search"
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "application/json",
               "Origin": "https://www.londonstockexchange.com",
               "Referer": "https://www.londonstockexchange.com/"}
    seen = set()
    for kw in TOKENIZATION_KEYWORDS[:8]:
        try:
            params = {"tidm": firm["ticker"], "searchTerm": kw,
                      "fromDate": START_DATE.replace("-",""), "toDate": END_DATE.replace("-",""),
                      "size": 50}
            r = requests.get(url, params=params, headers=headers, timeout=20)
            if r.status_code == 200:
                items = r.json().get("content", []) or r.json().get("results", [])
                for it in items:
                    head = it.get("headline", it.get("title",""))
                    date = str(it.get("date",""))[:10]
                    rid = str(it.get("rnsId", it.get("id","")))
                    if rid in seen or not _date_in_range(date):
                        continue
                    seen.add(rid)
                    rows.append(_base_row(firm, kw, date, it.get("regulatoryNewsType","RNS"),
                                          rid, "RNS",
                                          f"https://www.londonstockexchange.com/news-article/{firm['ticker']}/{rid}",
                                          head))
            time.sleep(0.25)
        except Exception as e:
            print(f"    RNS {firm['name']}/{kw}: {e}")
    return rows


# ── SIX ──────────────────────────────────────────────────────────────────────
def collect_six(firm) -> list:
    rows = []
    if not firm.get("isin"):
        return rows
    url = "https://www.six-group.com/sheldon/api/news/search"
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "application/json"}
    for kw in TOKENIZATION_KEYWORDS[:6]:
        try:
            params = {"query": kw, "isin": firm["isin"],
                      "from": START_DATE, "to": END_DATE, "language": "en"}
            r = requests.get(url, params=params, headers=headers, timeout=20)
            if r.status_code == 200:
                items = r.json().get("results", []) or r.json().get("news", [])
                for it in items:
                    head = it.get("title", it.get("headline",""))
                    date = str(it.get("date",""))[:10]
                    if not _date_in_range(date):
                        continue
                    rows.append(_base_row(firm, kw, date, it.get("type","AdHoc"),
                                          str(it.get("id","")), "SIX",
                                          it.get("url", it.get("link","")), head))
            time.sleep(0.25)
        except Exception as e:
            print(f"    SIX {firm['name']}/{kw}: {e}")
    return rows


# ── HKEX ─────────────────────────────────────────────────────────────────────
def collect_hkex(firm) -> list:
    rows = []
    url = "https://www1.hkexnews.hk/search/titleSearchServlet.do"
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "application/json"}
    seen = set()
    for kw in TOKENIZATION_KEYWORDS[:8]:
        try:
            params = {"sortDir": "0", "sortByOptions": "DateTime",
                      "category": "0", "market": "SEHK",
                      "stockId": firm["hkex_id"], "documentType": "-1",
                      "fromDate": START_DATE.replace("-",""), "toDate": END_DATE.replace("-",""),
                      "title": kw, "searchType": "1", "t1code": "-2", "t2Gcode": "-2", "t2code": "-2",
                      "lang": "EN", "rowRange": "100"}
            r = requests.get(url, params=params, headers=headers, timeout=20)
            if r.status_code == 200:
                try:
                    items = r.json().get("result", [])
                    if isinstance(items, str):
                        items = pd.read_json(items).to_dict("records")
                except Exception:
                    items = []
                for it in items:
                    head = it.get("TITLE", it.get("title",""))
                    date = str(it.get("DATE_TIME", it.get("date","")))[:10]
                    fid = str(it.get("FILE_ID", it.get("id","")))
                    if fid in seen or not _date_in_range(date):
                        continue
                    seen.add(fid)
                    rows.append(_base_row(firm, kw, date, it.get("TYPE","Disclosure"),
                                          fid, "HKEX",
                                          f"https://www1.hkexnews.hk{it.get('FILE_LINK','')}", head))
            time.sleep(0.4)
        except Exception as e:
            print(f"    HKEX {firm['name']}/{kw}: {e}")
    return rows


# ── TSE Japan (TDnet) ────────────────────────────────────────────────────────
def collect_tse(firm) -> list:
    """
    Japan TDnet disclosure search. TDnet has a JSON-ish search but coverage of
    English keyword matches is limited. Returns what it finds; low matches are
    expected. Fully automated (no manual step) per the requirement.
    """
    rows = []
    url = "https://webapi.yanoshin.jp/webapi/tdnet/list/{code}.json"
    try:
        r = requests.get(url.format(code=firm["ticker"]),
                         params={"limit": "300"},
                         headers={"User-Agent": "Mozilla/5.0 (research)"}, timeout=20)
        if r.status_code == 200:
            items = r.json().get("items", [])
            for wrap in items:
                it = wrap.get("Tdnet", {})
                head = it.get("title", "")
                date = str(it.get("pubdate", ""))[:10]
                if not _date_in_range(date):
                    continue
                if not any(k.lower() in head.lower() for k in TOKENIZATION_KEYWORDS):
                    continue
                rows.append(_base_row(firm, "match", date, "TDnet",
                                      str(it.get("id","")), "TDnet",
                                      it.get("document_url",""), head))
    except Exception as e:
        print(f"    TDnet {firm['name']}: {e}")
    time.sleep(0.4)
    return rows


COLLECTORS = {
    "US": collect_edgar, "SGX": collect_sgx, "XETRA": collect_eqs,
    "LSE": collect_rns, "SIX": collect_six, "HKEX": collect_hkex, "TSE": collect_tse,
}


# ═══════════════════════════════════════════════════════════════════════════════
# SIMFIN FINANCIALS
# ═══════════════════════════════════════════════════════════════════════════════

def get_simfin_financials() -> pd.DataFrame:
    """
    Pull fundamentals from SimFin for all firms:
    EBITDA, total debt, total equity, market cap → for D/E, EV, size controls.
    Uses the simfin Python package (paid API key required).
    """
    try:
        import simfin as sf
    except ImportError:
        print("  simfin not installed — run: !pip install simfin")
        return pd.DataFrame()

    sf.set_api_key(SIMFIN_API_KEY)
    sf.set_data_dir("/content/simfin_data")

    print("\n  Loading SimFin datasets (income, balance sheet, derived)...")
    rows = []

    # Load bulk datasets for the markets SimFin covers
    markets = ["us"]   # SimFin coverage: US strongest; add others if your plan supports
    try:
        for market in markets:
            income = sf.load_income(variant="annual", market=market)
            balance = sf.load_balance(variant="annual", market=market)
            derived = sf.load_derived(variant="annual", market=market)
            # SimFin indexes by Ticker + Fiscal Year; take latest year per ticker
            for firm in _all_firms():
                tk = firm["ticker"]
                try:
                    if tk in income.index.get_level_values("Ticker"):
                        inc = income.xs(tk, level="Ticker").iloc[-1]
                        bal = balance.xs(tk, level="Ticker").iloc[-1]
                        rows.append({
                            "ticker": tk, "firm_name": firm["name"],
                            "revenue": inc.get("Revenue", np.nan),
                            "ebitda_proxy": inc.get("Operating Income (Loss)", np.nan),
                            "total_debt": bal.get("Total Liabilities", np.nan),
                            "total_equity": bal.get("Total Equity", np.nan),
                            "simfin_source": market,
                        })
                except Exception:
                    continue
    except Exception as e:
        print(f"  SimFin bulk load issue: {e}")

    df = pd.DataFrame(rows)
    if not df.empty:
        df["debt_equity_ratio"] = df["total_debt"] / df["total_equity"]
    return df


def get_yfinance_fundamentals() -> pd.DataFrame:
    """
    Fallback / supplement for non-US firms SimFin may not cover.
    Pulls market cap, EBITDA, debt/equity from yfinance .info.
    """
    import yfinance as yf
    rows = []
    for firm in _all_firms():
        yft = firm.get("yf")
        if not yft:
            continue
        try:
            info = yf.Ticker(yft).info
            rows.append({
                "ticker": firm["ticker"], "firm_name": firm["name"],
                "yf_ticker": yft,
                "market_cap": info.get("marketCap", np.nan),
                "ebitda": info.get("ebitda", np.nan),
                "total_debt": info.get("totalDebt", np.nan),
                "debt_to_equity": info.get("debtToEquity", np.nan),
                "enterprise_value": info.get("enterpriseValue", np.nan),
                "sector": info.get("sector", ""),
            })
        except Exception as e:
            print(f"    yfinance {firm['name']}: {e}")
        time.sleep(0.2)
    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════════════════════
# PRICE COVERAGE CHECK
# ═══════════════════════════════════════════════════════════════════════════════

def check_price_coverage() -> pd.DataFrame:
    """
    Verify each firm has sufficient price history for the event study
    (need >= 260 trading days before earliest event).
    """
    import yfinance as yf
    rows = []
    for firm in _all_firms():
        yft = firm.get("yf")
        if not yft:
            rows.append({"firm_name": firm["name"], "ticker": firm["ticker"],
                         "yf_ticker": "NONE", "n_days": 0, "coverage_ok": False})
            continue
        try:
            hist = yf.download(yft, start="2020-01-01", end=END_DATE, progress=False)
            n = len(hist)
            rows.append({"firm_name": firm["name"], "ticker": firm["ticker"],
                         "yf_ticker": yft, "n_days": n, "coverage_ok": n >= 500})
        except Exception as e:
            rows.append({"firm_name": firm["name"], "ticker": firm["ticker"],
                         "yf_ticker": yft, "n_days": 0, "coverage_ok": False})
        time.sleep(0.2)
    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

def _all_firms():
    out = []
    for market, firms in FIRMS.items():
        for f in firms:
            f = dict(f); f["exchange"] = market
            out.append(f)
    return out


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
# BLOCKER 1 FIX — CANARY CROSS-CHECKS
# ═══════════════════════════════════════════════════════════════════════════════
#
# Each market has at least one KNOWN tokenization event. If the collector for a
# market does not find its canary, the endpoint has almost certainly changed or is
# blocking us — a zero result is then an API FAILURE, not a genuine "no events".
# The diagnostic block distinguishes the two and tells you exactly which.
# ═══════════════════════════════════════════════════════════════════════════════

CANARY_EVENTS = {
    # market : { firm substring, expected event in this month (YYYY-MM), description }
    "US":    {"firm": "BlackRock",   "month": "2024-03", "desc": "BlackRock BUIDL tokenized Treasury fund (Mar 2024)"},
    "SGX":   {"firm": "DBS",         "month": "2024-10", "desc": "DBS Token Services launch (Oct 2024)"},
    "XETRA": {"firm": "Siemens",     "month": "2023-02", "desc": "Siemens EUR 300M digital bond (Feb 2023)"},
    "LSE":   {"firm": "HSBC",        "month": "2023-11", "desc": "HSBC Orion tokenized bond platform (Nov 2023)"},
    "SIX":   {"firm": "UBS",         "month": "2023",    "desc": "UBS tokenized fund under Project Guardian (2023)"},
    "HKEX":  {"firm": "Bank of China","month": "2023",   "desc": "BOC HK digital/tokenized bond (2023)"},
    "TSE":   {"firm": "Nomura",      "month": "2023",    "desc": "Nomura Laser Digital RWA tokenization (2023)"},
}


def check_canary(market: str, events_df: pd.DataFrame) -> dict:
    """
    Determine whether a market's collector is WORKING.
    Returns a status dict used by the diagnostic block.

    Logic:
      - If the market has >0 events AND the canary firm appears -> OK
      - If >0 events but canary firm absent -> PARTIAL (endpoint works, but may
        be missing material events; widen keywords or verify manually)
      - If 0 events -> LIKELY_API_FAILURE (do not treat as a finding)
    """
    canary = CANARY_EVENTS.get(market)
    if events_df.empty:
        sub = pd.DataFrame()
    else:
        sub = events_df[events_df["exchange"] == market]

    n = len(sub)
    status = {"market": market, "n_events": n, "canary_desc": canary["desc"] if canary else "",
              "canary_found": False, "verdict": "", "action": ""}

    if canary is None:
        status["verdict"] = "NO_CANARY_DEFINED"
        return status

    # Does the canary firm appear at all?
    if not sub.empty:
        firm_match = sub["firm_name"].str.contains(canary["firm"], case=False, na=False)
        # Optional month match (tightens confidence)
        month_match = sub["event_date"].dt.strftime("%Y-%m") == canary["month"] if len(canary["month"]) == 7 else \
                      sub["event_date"].dt.strftime("%Y") == canary["month"]
        status["canary_found"] = bool((firm_match).any())
        status["canary_month_found"] = bool((firm_match & month_match).any())

    if n == 0:
        status["verdict"] = "LIKELY_API_FAILURE"
        status["action"] = (f"Market '{market}' returned ZERO events. This is almost certainly an "
                            f"endpoint change or block — NOT a real absence of tokenization news. "
                            f"Manually verify the collector against: {canary['desc']}. "
                            f"Check the endpoint URL/params in collect_{market.lower()}().")
    elif not status["canary_found"]:
        status["verdict"] = "PARTIAL_CANARY_MISSING"
        status["action"] = (f"Market '{market}' returned {n} events but the canary "
                            f"({canary['firm']}) is missing. Endpoint works but recall may be low. "
                            f"Widen keywords or manually confirm: {canary['desc']}.")
    else:
        status["verdict"] = "OK"
        status["action"] = f"Canary found ({canary['firm']}). Endpoint healthy."

    return status


# ═══════════════════════════════════════════════════════════════════════════════
# AUTOMATED VERIFICATION PASS
# ═══════════════════════════════════════════════════════════════════════════════
#
# Replaces the purely manual review with an automated first-pass that:
#   1. Auto-verifies HIGH-confidence events whose headline strongly matches a real
#      tokenization action (verb + asset class present)
#   2. Auto-rejects obvious noise (keyword present but in a non-tokenization context,
#      e.g. generic "digital" / "technology" boilerplate, ESG reports)
#   3. Leaves genuinely ambiguous events flagged AMBIGUOUS for the (now much smaller)
#      human pass
#
# This is conservative: it only auto-verifies when confidence is high AND the action
# is unambiguous. Everything else is surfaced explicitly.
# ═══════════════════════════════════════════════════════════════════════════════

# Strong signal: an action verb co-located with a tokenization object
STRONG_ACTION = r"\b(issue[ds]?|launch(ed|es)?|tokeniz(e|ed|ing)|complet(e|ed)|priced|settl(e|ed)|mint(ed)?|deploy(ed)?)\b"
STRONG_OBJECT = r"\b(bond|note|fund|token|securit|deposit|receivabl|real estate|treasury|gold|asset)\b"

# Noise signals: keyword present but likely NOT a tokenization event
NOISE_PATTERNS = [
    r"\b(annual report|sustainability|esg|csr|diversity|appoint|resign|dividend declaration|agm|egm|proxy)\b",
    r"\bdigital (transformation|banking app|channel|marketing|workplace)\b",   # "digital" but not assets
    r"\bblockchain (conference|webinar|partnership announcement only)\b",
]


def auto_verify(events_df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply automated verification. Adds columns:
      auto_status  : VERIFIED | REJECTED | AMBIGUOUS
      auto_reason  : human-readable rationale
      verified     : True only for VERIFIED (feeds Code 2 directly)
    """
    if events_df.empty:
        return events_df

    df = events_df.copy()
    statuses, reasons = [], []

    for _, ev in df.iterrows():
        head = str(ev.get("headline", "")).lower()
        conf = ev.get("classification_conf", "LOW")
        asset = ev.get("asset_class", "Unclassified")
        atype = ev.get("announcement_type", "Unclassified")

        # 1. Noise check first
        is_noise = any(re.search(p, head, re.IGNORECASE) for p in NOISE_PATTERNS)
        has_token_root = bool(re.search(r"token|on-chain|on chain|digital (bond|securit|asset)|distributed ledger", head, re.IGNORECASE))

        if is_noise and not has_token_root:
            statuses.append("REJECTED")
            reasons.append("Noise pattern matched (boilerplate/ESG/governance), no tokenization action")
            continue

        # 2. Strong-signal auto-verify
        has_action = bool(re.search(STRONG_ACTION, head, re.IGNORECASE))
        has_object = bool(re.search(STRONG_OBJECT, head, re.IGNORECASE))

        if conf == "HIGH" and has_action and has_object and has_token_root:
            statuses.append("VERIFIED")
            reasons.append(f"HIGH conf + action+object+token root in headline ({atype}/{asset})")
            continue

        if conf == "HIGH" and has_token_root and asset != "Unclassified":
            statuses.append("VERIFIED")
            reasons.append(f"HIGH conf + token root + classified asset ({asset})")
            continue

        # 3. Everything else is ambiguous -> small human pass
        statuses.append("AMBIGUOUS")
        missing = []
        if conf != "HIGH": missing.append(f"confidence={conf}")
        if asset == "Unclassified": missing.append("asset_class missing")
        if atype == "Unclassified": missing.append("announcement_type missing")
        if not has_token_root: missing.append("no clear tokenization root in headline")
        reasons.append("Needs human check: " + ("; ".join(missing) if missing else "borderline signal"))

    df["auto_status"] = statuses
    df["auto_reason"] = reasons
    df["verified"] = df["auto_status"] == "VERIFIED"
    return df


# ═══════════════════════════════════════════════════════════════════════════════
# DIAGNOSTIC / VERDICT BLOCK  (your "executed code with error line + fix" request)
# ═══════════════════════════════════════════════════════════════════════════════

def print_diagnostics(events_df, simfin_df, yf_df, coverage_df, canary_results,
                      markets, pull_financials, check_prices):
    """
    Single consolidated verdict block printed at the end of every run.
    For each subsystem: STATUS, what's short, and the exact fix to apply.
    Designed so you can act immediately without reading stack traces.
    """
    L = []
    L.append("")
    L.append("╔" + "═"*68 + "╗")
    L.append("║" + "  TAVE CODE 1 — RUN DIAGNOSTICS & VERDICT".ljust(68) + "║")
    L.append("╚" + "═"*68 + "╝")

    problems = []   # actionable items
    oks = []

    # ── 1. Event collection per market (Blocker 1) ──
    L.append("\n[1] EVENT COLLECTION BY MARKET")
    L.append("    " + "-"*60)
    for m in markets:
        cr = canary_results.get(m, {})
        verdict = cr.get("verdict", "UNKNOWN")
        n = cr.get("n_events", 0)
        if verdict == "OK":
            L.append(f"    ✓ {m:<6} {n:>4} events | canary OK")
            oks.append(m)
        elif verdict == "PARTIAL_CANARY_MISSING":
            L.append(f"    ⚠ {m:<6} {n:>4} events | CANARY MISSING")
            L.append(f"            → {cr.get('action','')}")
            problems.append(f"[{m}] canary missing — {cr.get('action','')}")
        elif verdict == "LIKELY_API_FAILURE":
            L.append(f"    ✗ {m:<6} {n:>4} events | LIKELY API FAILURE")
            L.append(f"            → {cr.get('action','')}")
            problems.append(f"[{m}] API FAILURE — {cr.get('action','')}")
        else:
            L.append(f"    ? {m:<6} {n:>4} events | {verdict}")

    # ── 2. Classification quality ──
    L.append("\n[2] CLASSIFICATION QUALITY")
    L.append("    " + "-"*60)
    if events_df.empty:
        L.append("    ✗ No events to classify.")
        problems.append("[CLASSIFY] No events collected at all — fix collection first.")
    else:
        conf_counts = events_df["classification_conf"].value_counts().to_dict()
        L.append(f"    Confidence: {conf_counts}")
        unclassified_asset = (events_df["asset_class"] == "Unclassified").sum()
        unclassified_type = (events_df["announcement_type"] == "Unclassified").sum()
        L.append(f"    asset_class unclassified:        {unclassified_asset}/{len(events_df)}")
        L.append(f"    announcement_type unclassified:  {unclassified_type}/{len(events_df)}")
        if unclassified_asset > len(events_df) * 0.5:
            problems.append("[CLASSIFY] >50% asset_class unclassified — headlines too terse. "
                            "FIX: enable body-text fetch for EDGAR (see fetch_filing_body flag) "
                            "or widen ASSET_CLASS_PATTERNS.")

    # ── 3. Automated verification ──
    L.append("\n[3] AUTOMATED VERIFICATION PASS")
    L.append("    " + "-"*60)
    if events_df.empty or "auto_status" not in events_df.columns:
        L.append("    ✗ Verification not run (no events).")
    else:
        vc = events_df["auto_status"].value_counts().to_dict()
        L.append(f"    {vc}")
        n_verified = vc.get("VERIFIED", 0)
        n_ambig = vc.get("AMBIGUOUS", 0)
        L.append(f"    → {n_verified} events auto-VERIFIED (feed Code 2 directly)")
        L.append(f"    → {n_ambig} events AMBIGUOUS (human pass on 'Manual_Review' sheet)")
        if n_verified < 30:
            problems.append(f"[VERIFY] Only {n_verified} auto-verified events. Event studies want "
                            f"30+ for large-effect power. FIX: review the AMBIGUOUS rows in "
                            f"Manual_Review and set verified=TRUE where genuine, OR widen keywords/firms.")
        else:
            oks.append("verification")

    # ── 4. Financials (SimFin) ──
    if pull_financials:
        L.append("\n[4] FINANCIALS (SimFin + yfinance)")
        L.append("    " + "-"*60)
        if simfin_df is None or simfin_df.empty:
            L.append("    ⚠ SimFin returned no rows.")
            problems.append("[SIMFIN] No SimFin data. FIX: confirm os.environ['SIMFIN_KEY'] is set "
                            "and your plan covers these markets. US coverage is strongest; "
                            "non-US fundamentals will come from yfinance instead.")
        else:
            L.append(f"    ✓ SimFin rows: {len(simfin_df)}")
        if yf_df is None or yf_df.empty:
            L.append("    ⚠ yfinance fundamentals empty.")
            problems.append("[YF-FUND] yfinance .info returned nothing. FIX: usually transient — "
                            "re-run; if persistent, some non-US tickers lack .info fields.")
        else:
            missing_mcap = yf_df["market_cap"].isna().sum() if "market_cap" in yf_df else len(yf_df)
            L.append(f"    ✓ yfinance rows: {len(yf_df)} | market_cap missing: {missing_mcap}")
            if missing_mcap > len(yf_df) * 0.3:
                problems.append("[YF-FUND] >30% market_cap missing — regression size control will be "
                                "weak. FIX: backfill from SimFin or annual reports for affected firms.")

    # ── 5. Price coverage (Blocker 2 preview) ──
    if check_prices:
        L.append("\n[5] PRICE COVERAGE (for Code 2 event study)")
        L.append("    " + "-"*60)
        if coverage_df is None or coverage_df.empty:
            L.append("    ✗ No coverage data.")
            problems.append("[PRICES] Coverage check returned nothing. FIX: yfinance/network issue, re-run.")
        else:
            ok_n = int(coverage_df["coverage_ok"].sum())
            bad = coverage_df[~coverage_df["coverage_ok"]]
            L.append(f"    ✓ Adequate coverage: {ok_n}/{len(coverage_df)} firms")
            if not bad.empty:
                L.append(f"    ⚠ Insufficient price history:")
                for _, r in bad.iterrows():
                    L.append(f"        - {r['firm_name']} ({r['yf_ticker']}): {r['n_days']} days")
                problems.append(f"[PRICES] {len(bad)} firms lack adequate history (<500 days). "
                                f"FIX: check the yf ticker is correct, or drop the firm. "
                                f"Code 2 will skip these automatically.")

    # ── FINAL VERDICT ──
    L.append("\n" + "═"*70)
    if not problems:
        L.append("  ✅ VERDICT: ALL SYSTEMS GREEN — proceed to manual pass (if any) then Code 2.")
    else:
        L.append(f"  ⚠️  VERDICT: {len(problems)} ITEM(S) NEED ATTENTION")
        L.append("  " + "-"*66)
        for i, p in enumerate(problems, 1):
            # wrap long lines
            L.append(f"  {i}. {p}")
    L.append("═"*70)

    print("\n".join(L))
    return problems


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ═══════════════════════════════════════════════════════════════════════════════

def run_collection(markets=("US","SGX","XETRA","LSE","SIX","HKEX","TSE"),
                   pull_financials=True, check_prices=True):
    print("█"*70)
    print("  TAVE EVENT COLLECTION — CODE 1  (with canary + auto-verify + diagnostics)")
    print(f"  {sum(len(FIRMS[m]) for m in markets if m in FIRMS)} firms | {START_DATE} → {END_DATE}")
    print("█"*70)

    # --- 0. Environment: mount Drive (if needed) + create output folder ---
    print("\n── Environment setup ──")
    ensure_drive_and_dirs()

    # --- 1. Collect events ---
    all_events = []
    for market in markets:
        if market not in FIRMS:
            continue
        print(f"\n── {market} ──")
        collector = COLLECTORS[market]
        for firm in tqdm(FIRMS[market], desc=market):
            firm = dict(firm); firm["exchange"] = market
            try:
                evs = collector(firm)
                all_events.extend(evs)
            except Exception as e:
                print(f"    {firm['name']}: {e}")

    events_df = pd.DataFrame(all_events)
    if not events_df.empty:
        events_df["event_date"] = pd.to_datetime(events_df["event_date"], errors="coerce")
        events_df = events_df.dropna(subset=["event_date"])
        events_df = events_df.sort_values(["firm_name","event_date"]).reset_index(drop=True)
        events_df.insert(0, "event_id", [f"EVT_{i:04d}" for i in range(1, len(events_df)+1)])

    # --- 1b. CANARY CHECK (Blocker 1) ---
    canary_results = {}
    for market in markets:
        if market in FIRMS:
            canary_results[market] = check_canary(market, events_df)

    # --- 1c. AUTOMATED VERIFICATION PASS ---
    if not events_df.empty:
        events_df = auto_verify(events_df)

    # --- 2. Financials ---
    simfin_df = pd.DataFrame(); yf_df = pd.DataFrame()
    if pull_financials:
        print("\n── SimFin financials ──")
        simfin_df = get_simfin_financials()
        print("\n── yfinance fundamentals (supplement) ──")
        yf_df = get_yfinance_fundamentals()

    # --- 3. Price coverage (Blocker 2 lives mainly in Code 2; previewed here) ---
    coverage_df = pd.DataFrame()
    if check_prices:
        print("\n── Price coverage check ──")
        coverage_df = check_price_coverage()

    # --- 4. Export ---
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        if not events_df.empty:
            events_df.to_excel(w, sheet_name="Events_Classified", index=False)
            # Manual review now = AMBIGUOUS only (auto-verify shrank this dramatically)
            review = events_df[events_df.get("auto_status", "") == "AMBIGUOUS"]
            review.to_excel(w, sheet_name="Manual_Review", index=False)
            # Rejected events kept for audit transparency
            rejected = events_df[events_df.get("auto_status", "") == "REJECTED"]
            if not rejected.empty:
                rejected.to_excel(w, sheet_name="Auto_Rejected", index=False)
        else:
            pd.DataFrame([{"note": "No events collected — see diagnostics"}]).to_excel(
                w, sheet_name="Events_Classified", index=False)
        # Canary results sheet
        pd.DataFrame(list(canary_results.values())).to_excel(w, sheet_name="Canary_Check", index=False)
        if not simfin_df.empty:
            simfin_df.to_excel(w, sheet_name="SimFin_Financials", index=False)
        if not yf_df.empty:
            yf_df.to_excel(w, sheet_name="YF_Fundamentals", index=False)
        if not coverage_df.empty:
            coverage_df.to_excel(w, sheet_name="Price_Coverage", index=False)

    print(f"\n✓ EXPORTED: {OUTPUT_FILE}")

    # --- 5. DIAGNOSTIC VERDICT BLOCK ---
    problems = print_diagnostics(events_df, simfin_df, yf_df, coverage_df,
                                 canary_results, markets, pull_financials, check_prices)

    # --- 6. COPY-PASTE SUMMARY ---
    print_copypaste_summary(events_df, canary_results, simfin_df, yf_df, coverage_df, markets)

    return events_df, simfin_df, yf_df, coverage_df, canary_results, problems


def print_copypaste_summary(events_df, canary_results, simfin_df, yf_df, coverage_df, markets):
    """
    Plain-text summary block designed to be copied directly into a chat or notes
    for interpretation. Reports what was collected, the shape of the data, and the
    headline numbers — no jargon, no stack traces.
    """
    S = []
    S.append("\n\n" + "┌" + "─"*68 + "┐")
    S.append("│" + "  COPY-PASTE SUMMARY — CODE 1 (COLLECTION)".ljust(68) + "│")
    S.append("│" + "  Paste this block for interpretation guidance".ljust(68) + "│")
    S.append("└" + "─"*68 + "┘")

    if events_df is None or events_df.empty:
        S.append("\nNo events were collected. See the diagnostics above for the cause")
        S.append("(most likely an API endpoint failure in one or more markets).")
        print("\n".join(S)); return

    n_total = len(events_df)
    n_verified = int((events_df.get("auto_status", "") == "VERIFIED").sum()) if "auto_status" in events_df else 0
    n_ambig = int((events_df.get("auto_status", "") == "AMBIGUOUS").sum()) if "auto_status" in events_df else 0
    n_rejected = int((events_df.get("auto_status", "") == "REJECTED").sum()) if "auto_status" in events_df else 0
    n_firms = events_df["firm_name"].nunique()
    date_lo = events_df["event_date"].min()
    date_hi = events_df["event_date"].max()

    S.append(f"\nWHAT WAS COLLECTED")
    S.append(f"  Candidate events:        {n_total}")
    S.append(f"  Auto-verified (usable):  {n_verified}")
    S.append(f"  Ambiguous (human pass):  {n_ambig}")
    S.append(f"  Auto-rejected (noise):   {n_rejected}")
    S.append(f"  Distinct firms:          {n_firms}")
    S.append(f"  Date span:               {str(date_lo)[:10]} to {str(date_hi)[:10]}")

    S.append(f"\nEVENTS BY MARKET (verified / total)")
    for m in markets:
        sub = events_df[events_df["exchange"] == m]
        v = int((sub.get("auto_status", "") == "VERIFIED").sum()) if "auto_status" in sub else 0
        cr = canary_results.get(m, {})
        flag = "" if cr.get("verdict") == "OK" else f"  [{cr.get('verdict','')}]"
        S.append(f"  {m:<6} {v:>3} / {len(sub):<3}{flag}")

    S.append(f"\nVERIFIED EVENTS BY ASSET CLASS")
    ver = events_df[events_df.get("auto_status", "") == "VERIFIED"] if "auto_status" in events_df else events_df
    if not ver.empty:
        for ac, n in ver["asset_class"].value_counts().items():
            S.append(f"  {ac:<18} {n}")

    S.append(f"\nVERIFIED EVENTS BY ANNOUNCEMENT TYPE")
    if not ver.empty:
        for at, n in ver["announcement_type"].value_counts().items():
            S.append(f"  {at:<18} {n}")

    S.append(f"\nVERIFIED EVENTS BY JURISDICTION")
    if not ver.empty:
        for j, n in ver["jurisdiction"].value_counts().items():
            S.append(f"  {j:<14} {n}")

    audit_n = int(ver["audit_disclosed"].sum()) if not ver.empty else 0
    pct_n = int(ver["pct_tokenized"].notna().sum()) if not ver.empty else 0
    S.append(f"\nDISCLOSURE SIGNALS (verified events)")
    S.append(f"  Smart-contract audit disclosed:  {audit_n}/{len(ver)}")
    S.append(f"  % balance sheet tokenized given: {pct_n}/{len(ver)}")

    # Data readiness for Code 2
    cov_ok = int(coverage_df["coverage_ok"].sum()) if coverage_df is not None and not coverage_df.empty else "n/a"
    cov_tot = len(coverage_df) if coverage_df is not None and not coverage_df.empty else "n/a"
    S.append(f"\nDATA READINESS FOR CODE 2")
    S.append(f"  Firms with adequate price history: {cov_ok}/{cov_tot}")
    S.append(f"  SimFin financial rows:             {0 if simfin_df is None or simfin_df.empty else len(simfin_df)}")
    S.append(f"  yfinance fundamental rows:         {0 if yf_df is None or yf_df.empty else len(yf_df)}")

    S.append(f"\nDELIVERABLE")
    S.append(f"  File: {OUTPUT_FILE}")
    S.append(f"  Sheets: Events_Classified, Manual_Review, Auto_Rejected,")
    S.append(f"          Canary_Check, SimFin_Financials, YF_Fundamentals, Price_Coverage")

    S.append(f"\nNEXT STEP")
    if n_verified >= 30:
        S.append(f"  {n_verified} verified events is adequate. Skim Manual_Review to promote any")
        S.append(f"  genuine events, then run Code 2.")
    else:
        S.append(f"  Only {n_verified} verified events. Promote genuine events from Manual_Review")
        S.append(f"  (set verified=TRUE), or widen keywords/firms, before running Code 2.")

    S.append("\n" + "─"*70)
    S.append("END SUMMARY — copy from the top border to here.")
    S.append("─"*70)
    print("\n".join(S))


if __name__ == "__main__":
    # Recommended first run — US only, to validate EDGAR:
    # run_collection(markets=("US",), pull_financials=False, check_prices=False)

    # Full run:
    run_collection()
