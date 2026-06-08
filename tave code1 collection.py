"""
═══════════════════════════════════════════════════════════════════════════════
  TAVE EVENT STUDY — CODE 1 of 2
  EVENT COLLECTION, CLASSIFICATION & DATA ASSEMBLY
═══════════════════════════════════════════════════════════════════════════════

  Purpose:
    1. Collect tokenization announcements. Default markets: US (SEC EDGAR),
       Singapore (SGX), Switzerland (SIX) — three diversified jurisdictions.
       (Firm lists for XETRA / LSE / HKEX / TSE remain defined and can be
       re-enabled, but are deferred to future research: their feeds require
       commercial licensing or have fragile/blocked endpoints.)
    2. Auto-classify each event: type, asset class, jurisdiction,
       audit disclosure, % balance sheet tokenized
    3. De-duplicate near-identical filings and score each event for impact
       (HIGH / MEDIUM / LOW) so the cleanest events can be isolated
    4. Auto-verify clear events; flag only ambiguous ones for a short human pass
    5. Pull financials from SimFin (optional) + prices/fundamentals from yfinance
    6. Export ONE master Excel file to Google Drive

  Output: tave_master_database.xlsx with sheets:
    - Events_Classified   (all events + classification + impact score/tier)
    - Manual_Review       (only AMBIGUOUS events needing a human glance)
    - Auto_Rejected       (noise filtered out, kept for transparency)
    - Canary_Check        (per-market endpoint health)
    - SimFin_Financials / YF_Fundamentals / Price_Coverage

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


# SimFin key — accepts either secret name (SIMFIN_API_KEY or SIMFIN_KEY),
# from Colab secrets or an environment variable.
SIMFIN_API_KEY = (_get_secret("SIMFIN_API_KEY")
                  or _get_secret("SIMFIN_KEY")
                  or "PUT_YOUR_KEY_HERE")

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
        print("  ⚠ SimFin key not found. Add a Colab secret named SIMFIN_API_KEY")
        print("    (or SIMFIN_KEY) via the key icon, and toggle notebook access ON.")
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
        {"name": "DBS Group Holdings",  "ticker": "D05",  "yf": "D05.SI",  "stock_id": "D05",  "jurisdiction": "Singapore", "newsroom_url": "https://www.dbs.com/newsroom/archive.page"},
        {"name": "Singapore Exchange",  "ticker": "S68",  "yf": "S68.SI",  "stock_id": "S68",  "jurisdiction": "Singapore", "newsroom_url": "https://www.sgxgroup.com/media-centre/news-releases"},
        {"name": "OCBC Bank",           "ticker": "O39",  "yf": "O39.SI",  "stock_id": "O39",  "jurisdiction": "Singapore", "newsroom_url": "https://www.ocbc.com/group/media/index.page"},
        {"name": "United Overseas Bank","ticker": "U11",  "yf": "U11.SI",  "stock_id": "U11",  "jurisdiction": "Singapore", "newsroom_url": "https://www.uobgroup.com/web-resources/uobgroup/press/news.html"},
        {"name": "CapitaLand Invest",   "ticker": "9CI",  "yf": "9CI.SI",  "stock_id": "9CI",  "jurisdiction": "Singapore", "newsroom_url": "https://www.capitaland.com/en/about-capitaland/newsroom.html"},
        {"name": "Keppel Ltd",          "ticker": "BN4",  "yf": "BN4.SI",  "stock_id": "BN4",  "jurisdiction": "Singapore", "newsroom_url": "https://www.keppel.com/news-and-resources/"},
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
        {"name": "UBS Group",           "ticker": "UBSG", "yf": "UBSG.SW", "isin": "CH0244767585", "jurisdiction": "Switzerland", "newsroom_url": "https://www.ubs.com/global/en/media/display-page-ndp/en-20241101-media-releases.html"},
        {"name": "Julius Baer",         "ticker": "BAER", "yf": "BAER.SW", "isin": "CH0102484968", "jurisdiction": "Switzerland", "newsroom_url": "https://www.juliusbaer.com/en/news/"},
        {"name": "SIX Swiss (Pfd proxy)","ticker":"SIXN", "yf": None,      "isin": "CH0362432707", "jurisdiction": "Switzerland", "newsroom_url": "https://www.six-group.com/en/newsroom/media-releases.html"},
        {"name": "Zurich Insurance",    "ticker": "ZURN", "yf": "ZURN.SW", "isin": "CH0011075394", "jurisdiction": "Switzerland", "newsroom_url": "https://www.zurich.com/media/news-releases"},
        {"name": "Partners Group",      "ticker": "PGHN", "yf": "PGHN.SW", "isin": "CH0024608827", "jurisdiction": "Switzerland", "newsroom_url": "https://www.partnersgroup.com/en/news-views/"},
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
        # keep a short body snippet so auto_verify can inspect real content
        "body_snippet":       (body or "")[:500],
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
def _fetch_edgar_doc_text(cik_num, accession, max_chars=8000):
    """
    Fetch the primary document text of an EDGAR filing so we can classify on real
    content rather than just metadata. Returns a lowercased text snippet or "".
    Conservative: one request, short timeout, capped length.
    """
    try:
        # The filing index lists documents; the primary doc is usually the first .htm
        base = f"https://www.sec.gov/Archives/edgar/data/{cik_num}/{accession.replace('-','')}"
        idx_url = f"{base}/{accession}-index.htm"
        r = requests.get(idx_url, headers={"User-Agent": SEC_USER_AGENT}, timeout=15)
        if r.status_code != 200:
            return ""
        # Find the primary document link (first .htm that isn't the index itself)
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(r.text, "lxml")
        doc_link = None
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if href.endswith(".htm") and "index" not in href.lower():
                doc_link = href if href.startswith("http") else "https://www.sec.gov" + href
                break
        if not doc_link:
            return ""
        time.sleep(0.12)
        rd = requests.get(doc_link, headers={"User-Agent": SEC_USER_AGENT}, timeout=15)
        if rd.status_code != 200:
            return ""
        txt = BeautifulSoup(rd.text, "lxml").get_text(separator=" ", strip=True)
        return txt[:max_chars].lower()
    except Exception:
        return ""


def collect_edgar(firm, fetch_body=True) -> list:
    """
    Collect material tokenization announcements from SEC EDGAR.

    Two important changes from the naive version:
      1. ONLY 8-K and 6-K (material event filings). 10-K/10-Q/20-F are annual/
         quarterly reports that mention 'blockchain' as boilerplate — they flooded
         the sample with thousands of non-events. Material announcements are 8-Ks.
      2. Classify primarily from the matched keyword phrase (which already encodes
         the asset class, e.g. "tokenized bond" -> Bond). Only fetch the filing body
         for the small number of filings where the keyword is generic AND the
         classification is still incomplete. This cuts runtime dramatically because
         we avoid thousands of document downloads.
    """
    rows = []
    endpoint = "https://efts.sec.gov/LATEST/search-index"
    headers = {"User-Agent": SEC_USER_AGENT}

    # Specific keywords carry their own asset-class signal; generic ones don't.
    specific_keywords = [
        '"tokenized fund"', '"tokenised fund"', '"tokenized bond"',
        '"tokenized treasury"', '"tokenized deposit"', '"digital bond"',
        '"tokenized securities"', '"tokenised securities"', '"on-chain fund"',
    ]
    generic_keywords = [
        '"asset tokenization"', '"real-world asset"', '"real world asset tokeniz"',
        '"tokenization platform"', '"tokenization of"',
    ]
    edgar_keywords = specific_keywords + generic_keywords

    # First pass: gather unique filings + the keyword that matched (no body fetch yet)
    candidates = {}   # accession -> {date, form, kw, matched_specific}
    for kw in edgar_keywords:
        try:
            params = {
                "q": kw, "dateRange": "custom",
                "startdt": START_DATE, "enddt": END_DATE,
                "entity": firm["name"],
                "forms": "8-K,6-K",
            }
            r = requests.get(endpoint, params=params, headers=headers, timeout=20)
            if r.status_code != 200:
                time.sleep(0.2); continue
            hits = r.json().get("hits", {}).get("hits", [])
            for hit in hits:
                src = hit.get("_source", {})
                acc = hit.get("_id", "").split(":")[0]
                date = src.get("file_date", "")
                if not _date_in_range(date):
                    continue
                # Keep the most specific keyword that matched this filing
                if acc not in candidates or (kw in specific_keywords and not candidates[acc]["matched_specific"]):
                    candidates[acc] = {
                        "date": date, "form": src.get("form_type", "8-K"),
                        "kw": kw.strip('"'), "matched_specific": kw in specific_keywords,
                    }
            time.sleep(0.15)
        except Exception as e:
            print(f"    EDGAR {firm['name']}/{kw}: {e}")

    # Second pass: build rows. Fetch body ONLY when the keyword was generic
    # (a specific keyword like "tokenized bond" already classifies the asset).
    cik_num = firm["cik"].lstrip("0")
    n_fetched = 0
    for acc, info in candidates.items():
        url = f"https://www.sec.gov/Archives/edgar/data/{cik_num}/{acc.replace('-','')}/{acc}-index.htm"
        pseudo_headline = f"{firm['name']} {info['form']} {info['kw']}"

        body = ""
        prelim = classify_event(pseudo_headline)

        if info["matched_specific"]:
            # Specific keyword already gives us the asset class. An 8-K filing of a
            # specific tokenization phrase is a material event; default the type to
            # "New Program" if not otherwise determinable. No body fetch needed.
            row = _base_row(firm, info["kw"], info["date"], info["form"], acc,
                            "EDGAR", url, pseudo_headline, body="")
            if row["announcement_type"] == "Unclassified":
                row["announcement_type"] = "New Program"
                # bump confidence: specific keyword + material 8-K = HIGH
                if row["asset_class"] != "Unclassified":
                    row["classification_conf"] = "HIGH"
            rows.append(row)
        else:
            # Generic keyword — fetch the body to determine asset class & type
            need_body = (prelim["asset_class"] == "Unclassified" or
                         prelim["announcement_type"] == "Unclassified")
            if fetch_body and need_body:
                body = _fetch_edgar_doc_text(cik_num, acc)
                n_fetched += 1
            rows.append(_base_row(
                firm, info["kw"], info["date"], info["form"], acc,
                "EDGAR", url, pseudo_headline, body=body,
            ))
    return rows


# ── SGX ──────────────────────────────────────────────────────────────────────
def collect_sgx(firm) -> list:
    """
    SGX company announcements.

    Fixes from the version that missed the DBS canary:
      1. SGX's JSON sometimes arrives with junk prefix chars (e.g. '{}&&') — strip
         before parsing.
      2. The keyword filter was too strict on a terse headline. SGX exposes a
         'title'/'headline' and sometimes a longer description; check BOTH, and
         broaden the keyword set so material events (DBS Token Services) are caught.
      3. Pull more pages and a wider date sort so older 2021-2024 events aren't
         truncated by pagination.
    """
    rows = []
    if firm.get("us_listed"):   # e.g. Sea Ltd files with SEC
        return collect_edgar(firm)
    if not firm.get("stock_id"):
        return rows

    url = "https://api.sgx.com/securities/v1.1/announcements"
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "application/json, text/plain"}

    def _parse_sgx_json(text):
        # Strip leading junk like "{}&&" that SGX sometimes prepends
        t = text.strip()
        if t.startswith("{}&&"):
            t = t[4:]
        # Some variants prefix with ")]}'," or similar — strip up to first '{'
        brace = t.find("{")
        if brace > 0:
            t = t[brace:]
        try:
            return json.loads(t)
        except Exception:
            return {}

    try:
        # Paginate to capture the full 2021-2025 window
        seen = set()
        for page in range(0, 6):   # up to 6 pages * 250 = 1500 announcements
            params = {"pagestart": str(page), "pagesize": "250",
                      "params": "id,announcementId,securityName,headline,title,"
                                "category,subCategory,announcedOn,url,documentLink"}
            # SGX filters by security via the 'value' on a related endpoint; the
            # general endpoint returns all, so we filter by securityName ourselves.
            r = requests.get(url, params=params, headers=headers, timeout=25)
            if r.status_code != 200:
                break
            data = r.json() if r.headers.get("content-type","").startswith("application/json") else _parse_sgx_json(r.text)
            anns = (data.get("data", {}) or {}).get("announcements", []) or data.get("data", []) or []
            if not anns:
                break

            for a in anns:
                # Match this firm by name (SGX general feed isn't always per-security)
                sec_name = str(a.get("securityName", a.get("security", ""))).lower()
                firm_key = firm["name"].lower().split()[0]   # e.g. 'dbs'
                # If the feed is per-security (stockId honored) sec_name may be blank — accept then
                if sec_name and firm_key not in sec_name and firm["ticker"].lower() not in sec_name:
                    continue

                head = str(a.get("headline", a.get("title", "")))
                desc = str(a.get("subCategory", "")) + " " + str(a.get("category", ""))
                combined = f"{head} {desc}".lower()

                date_raw = str(a.get("announcedOn", a.get("announcementDate", "")))[:10]
                if not _date_in_range(date_raw):
                    continue

                # Broader keyword match across headline + category
                if not any(k.lower() in combined for k in TOKENIZATION_KEYWORDS):
                    continue

                ann_id = str(a.get("announcementId", a.get("id", head[:20] + date_raw)))
                if ann_id in seen:
                    continue
                seen.add(ann_id)

                link = a.get("url", a.get("documentLink", ""))
                rows.append(_base_row(
                    firm, "match", date_raw, a.get("category", "SGX-Ann"),
                    ann_id, "SGX", link, head,
                ))
            time.sleep(0.4)
    except Exception as e:
        print(f"    SGX {firm['name']}: {e}")
    time.sleep(0.3)

    # Supplement with the company newsroom — catches product-launch announcements
    # (e.g. DBS Token Services) that are NOT filed as SGX regulatory disclosures.
    rows.extend(collect_newsroom(firm))
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


# ── COMPANY NEWSROOM (supplements exchange feeds) ────────────────────────────
def collect_newsroom(firm) -> list:
    """
    Scrape a company's own newsroom / press-release page for tokenization news.

    WHY THIS EXISTS: Many tokenization announcements are corporate press releases,
    NOT regulatory filings. Exchange feeds (SGX/SIX regulatory news) are dominated
    by mandatory disclosures (results, dividends, AGMs) and often MISS product
    launches like 'DBS Token Services'. The company newsroom is where those live.
    This is the key source for catching events the exchange feed misses.

    Requires firm['newsroom_url']. Returns [] if not provided or unreachable.
    """
    rows = []
    url = firm.get("newsroom_url")
    if not url:
        return rows
    from bs4 import BeautifulSoup
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "text/html"}

    try:
        r = requests.get(url, headers=headers, timeout=20)
        if r.status_code != 200:
            return rows
        soup = BeautifulSoup(r.text, "lxml")
        # Collect candidate news blocks: links + nearby text
        seen = set()
        for a in soup.find_all("a", href=True):
            title = a.get_text(separator=" ", strip=True)
            if len(title) < 15 or len(title) > 300:
                continue
            low = title.lower()
            if not any(k.lower() in low for k in TOKENIZATION_KEYWORDS):
                continue
            # Try to find a date near this link (parent text)
            parent_txt = a.find_parent().get_text(separator=" ", strip=True) if a.find_parent() else title
            m = re.search(r"(\d{1,2}\s+\w+\s+\d{4}|\d{4}-\d{2}-\d{2}|\w+\s+\d{1,2},\s+\d{4})", parent_txt)
            date = ""
            if m:
                for fmt in ("%Y-%m-%d", "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y"):
                    try:
                        date = datetime.strptime(m.group(1), fmt).strftime("%Y-%m-%d"); break
                    except ValueError:
                        continue
            # If no date found, skip (event study needs a date) — but keep for manual review
            if not date or not _date_in_range(date):
                continue
            key = title[:60]
            if key in seen:
                continue
            seen.add(key)
            link = a["href"]
            if link.startswith("/"):
                from urllib.parse import urlparse
                base = f"{urlparse(url).scheme}://{urlparse(url).netloc}"
                link = base + link
            rows.append(_base_row(firm, "match", date, "Newsroom",
                                  key, "Newsroom", link, title))
    except Exception as e:
        print(f"    Newsroom {firm['name']}: {e}")
    time.sleep(0.4)
    return rows


# ── SIX (Switzerland) ────────────────────────────────────────────────────────
def collect_six(firm) -> list:
    """
    SIX Swiss Exchange regulatory news.

    The JSON API requires a commercial data licence (it returned zero for us).
    The robust free path is the issuer's regulatory disclosure page on six-group.com,
    which lists ad-hoc announcements as HTML. We fetch the issuer's news listing and
    keyword-filter. If SIX's structure blocks us, we fall back to the company's own
    investor-relations newsroom (most Swiss blue-chips publish ad-hoc news there too).
    """
    rows = []
    from bs4 import BeautifulSoup
    headers = {"User-Agent": "Mozilla/5.0 (research)", "Accept": "text/html"}

    # SIX issuer explorer regulatory-news listing (by ISIN)
    candidate_urls = []
    if firm.get("isin"):
        candidate_urls.append(
            f"https://www.six-group.com/en/market-data/shares/share-explorer/regulatory-news.html?isin={firm['isin']}")
    # Company IR newsroom fallback (firm-specific override optional)
    if firm.get("ir_news_url"):
        candidate_urls.append(firm["ir_news_url"])

    for url in candidate_urls:
        try:
            r = requests.get(url, headers=headers, timeout=20)
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.text, "lxml")
            # Generic extraction: find list items / rows containing a date and a headline
            text_blocks = soup.find_all(["article", "li", "tr", "div"], limit=400)
            seen = set()
            for blk in text_blocks:
                txt = blk.get_text(separator=" ", strip=True)
                if len(txt) < 20 or len(txt) > 400:
                    continue
                low = txt.lower()
                if not any(k.lower() in low for k in TOKENIZATION_KEYWORDS):
                    continue
                # Extract a date if present
                m = re.search(r"(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{4}-\d{2}-\d{2})", txt)
                date = ""
                if m:
                    raw = m.group(1)
                    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%y"):
                        try:
                            date = datetime.strptime(raw, fmt).strftime("%Y-%m-%d"); break
                        except ValueError:
                            continue
                if not date or not _date_in_range(date):
                    continue
                key = txt[:60]
                if key in seen:
                    continue
                seen.add(key)
                link_el = blk.find("a", href=True)
                link = link_el["href"] if link_el else url
                if link.startswith("/"):
                    link = "https://www.six-group.com" + link
                rows.append(_base_row(firm, "match", date, "SIX-AdHoc",
                                      key, "SIX", link, txt[:160]))
            if rows:
                break   # got data from this URL, don't try fallback
            time.sleep(0.3)
        except Exception as e:
            print(f"    SIX {firm['name']}: {e}")
    time.sleep(0.3)

    # Supplement with the company newsroom (Swiss blue-chips publish ad-hoc news there)
    rows.extend(collect_newsroom(firm))
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

    Decisions use the combined headline + body snippet (the real filing text),
    and the classification result — not the thin metadata headline alone.
    """
    if events_df.empty:
        return events_df

    df = events_df.copy()
    statuses, reasons = [], []

    for _, ev in df.iterrows():
        # Use BOTH headline and the fetched body snippet — this is the key fix.
        text = (str(ev.get("headline", "")) + " " + str(ev.get("body_snippet", ""))).lower()
        conf = ev.get("classification_conf", "LOW")
        asset = ev.get("asset_class", "Unclassified")
        atype = ev.get("announcement_type", "Unclassified")

        # 1. Noise check
        is_noise = any(re.search(p, text, re.IGNORECASE) for p in NOISE_PATTERNS)
        has_token_root = bool(re.search(
            r"token|on-chain|on chain|digital (bond|securit|asset)|distributed ledger|real[- ]world asset",
            text, re.IGNORECASE))

        if is_noise and not has_token_root:
            statuses.append("REJECTED")
            reasons.append("Noise pattern matched (boilerplate/ESG/governance), no tokenization action")
            continue

        has_action = bool(re.search(STRONG_ACTION, text, re.IGNORECASE))
        has_object = bool(re.search(STRONG_OBJECT, text, re.IGNORECASE))

        # 2. Strong-signal auto-verify (classification HIGH + real content signals)
        if conf == "HIGH" and has_token_root and (has_action or asset != "Unclassified"):
            statuses.append("VERIFIED")
            reasons.append(f"HIGH conf + token root + action/asset in content ({atype}/{asset})")
            continue

        # 3. Medium confidence but clear action+object+token root in the body = verify
        if conf in ("HIGH", "MEDIUM") and has_action and has_object and has_token_root:
            statuses.append("VERIFIED")
            reasons.append(f"action+object+token root in filing text ({atype}/{asset})")
            continue

        # 4. Otherwise ambiguous -> human pass
        statuses.append("AMBIGUOUS")
        missing = []
        if conf == "LOW": missing.append(f"confidence={conf}")
        if asset == "Unclassified": missing.append("asset_class missing")
        if atype == "Unclassified": missing.append("announcement_type missing")
        if not has_token_root: missing.append("no clear tokenization signal in content")
        reasons.append("Needs human check: " + ("; ".join(missing) if missing else "borderline signal"))

    df["auto_status"] = statuses
    df["auto_reason"] = reasons
    df["verified"] = df["auto_status"] == "VERIFIED"
    return df


# ═══════════════════════════════════════════════════════════════════════════════
# DEDUPLICATION + HIGH-IMPACT SCORING  (US tightening — request A)
# ═══════════════════════════════════════════════════════════════════════════════
#
# Two problems with raw collection:
#  1. The SAME real-world announcement often appears as multiple filings within a
#     few days (an 8-K, then an amended 8-K, then a related 6-K). Counting each as
#     a separate "event" inflates the sample and violates event-study independence.
#  2. Not every filing that mentions tokenization is a MATERIAL event. We want a
#     score that surfaces genuine announcements over passing references.
# ═══════════════════════════════════════════════════════════════════════════════

def deduplicate_events(events_df: pd.DataFrame, window_days: int = 5) -> pd.DataFrame:
    """
    Collapse near-duplicate filings into single events.
    Rule: same firm + same asset_class within `window_days` = one event.
    Keeps the highest-impact row in each cluster (see impact score).
    Adds a 'dup_group_size' column recording how many filings collapsed.
    """
    if events_df.empty:
        return events_df

    df = events_df.sort_values(["firm_name", "event_date"]).copy()
    df["_keep"] = True
    df["dup_group_size"] = 1

    # Ensure impact score exists for tie-breaking
    if "impact_score" not in df.columns:
        df = score_impact(df)

    for firm in df["firm_name"].unique():
        fdf = df[df["firm_name"] == firm]
        # cluster within asset class
        for ac in fdf["asset_class"].unique():
            sub = fdf[fdf["asset_class"] == ac].sort_values("event_date")
            if len(sub) < 2:
                continue
            cluster = []
            last_date = None
            for idx, row in sub.iterrows():
                d = row["event_date"]
                if last_date is not None and (d - last_date).days <= window_days:
                    cluster.append(idx)
                else:
                    if len(cluster) > 1:
                        _resolve_cluster(df, cluster)
                    cluster = [idx]
                last_date = d
            if len(cluster) > 1:
                _resolve_cluster(df, cluster)

    out = df[df["_keep"]].drop(columns=["_keep"]).reset_index(drop=True)
    return out


def _resolve_cluster(df, idxs):
    """Keep the highest-impact row in a duplicate cluster; mark the rest dropped."""
    sub = df.loc[idxs]
    keep_idx = sub["impact_score"].idxmax()
    for i in idxs:
        if i != keep_idx:
            df.at[i, "_keep"] = False
    df.at[keep_idx, "dup_group_size"] = len(idxs)


# High-impact signal vocabulary
IMPACT_STRONG = r"\b(launch\w*|issu\w+|complet\w+|first|inaugural|priced|settl\w+|live|go[- ]live|deploy\w*|mint\w*|million|billion|partnership with|in collaboration)\b"
IMPACT_MEDIUM = r"\b(expand\w*|pilot\w*|plan\w*|intend\w*|explor\w*|develop\w*|announc\w+)\b"
IMPACT_WEAK   = r"\b(consider\w*|potential\w*|may |could |evaluat\w*|research\w*|study\w*)\b"


def score_impact(events_df: pd.DataFrame) -> pd.DataFrame:
    """
    Assign each event an impact_score [0-100] and an impact_tier.
    Logic combines:
      - filing form (8-K material event > 6-K > others)
      - classification confidence (HIGH > MEDIUM > LOW)
      - asset class identified (named asset > unclassified)
      - language strength in headline+body (launch/issue/$amount > explore/consider)
      - whether an amount or % was disclosed (concrete > vague)
      - audit disclosed (extra credibility)
    """
    if events_df.empty:
        events_df["impact_score"] = []
        events_df["impact_tier"] = []
        return events_df

    df = events_df.copy()
    scores = []
    for _, ev in df.iterrows():
        s = 0
        text = (str(ev.get("headline", "")) + " " + str(ev.get("body_snippet", ""))).lower()

        # Form type
        form = str(ev.get("form_type", "")).upper()
        if "8-K" in form: s += 20
        elif "6-K" in form: s += 15
        else: s += 5

        # Confidence
        conf = ev.get("classification_conf", "LOW")
        s += {"HIGH": 25, "MEDIUM": 15, "LOW": 5}.get(conf, 5)

        # Asset class identified
        if ev.get("asset_class", "Unclassified") != "Unclassified": s += 15

        # Announcement type strength
        atype = ev.get("announcement_type", "Unclassified")
        s += {"Live Launch": 15, "Expansion": 10, "New Program": 8, "Pilot": 5}.get(atype, 0)

        # Language strength
        if re.search(IMPACT_STRONG, text): s += 15
        elif re.search(IMPACT_MEDIUM, text): s += 8
        if re.search(IMPACT_WEAK, text): s -= 5

        # Concrete disclosures
        if pd.notna(ev.get("amount_disclosed")) and ev.get("amount_disclosed"): s += 8
        if pd.notna(ev.get("pct_tokenized")) and ev.get("pct_tokenized"): s += 4
        if ev.get("audit_disclosed"): s += 5

        scores.append(max(0, min(100, s)))

    df["impact_score"] = scores
    # Thresholds calibrated so verified events spread across tiers rather than
    # all landing in HIGH. HIGH now requires genuinely strong signals (named asset
    # + concrete amount/% + launch/issue language), MEDIUM is the typical material
    # announcement, LOW is thin. This gives the robustness regression a real subset.
    df["impact_tier"] = pd.cut(df["impact_score"], bins=[-1, 60, 80, 101],
                               labels=["LOW", "MEDIUM", "HIGH"])
    return df



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

def _load_cached_market(market):
    """
    Return previously-collected events for a market from the saved Excel file,
    or None if absent. Lets us skip re-scraping a market we already have.
    """
    if not os.path.isfile(OUTPUT_FILE):
        return None
    try:
        prev = pd.read_excel(OUTPUT_FILE, sheet_name="Events_Classified")
        prev = prev[prev["exchange"] == market].copy()
        if prev.empty:
            return None
        prev["event_date"] = pd.to_datetime(prev["event_date"], errors="coerce")
        return prev
    except Exception:
        return None


def _merge_market_keep_best(fresh_df, market):
    """
    'Never downgrade' guard for the living document.
    Compare a fresh scrape of one market against what's already cached. If the
    fresh scrape returned FEWER events than the cache (likely an API hiccup),
    keep the cached version and warn. Otherwise use the fresh one.
    Returns (df_to_use, note).
    """
    cached = _load_cached_market(market)
    n_fresh = 0 if fresh_df is None else len(fresh_df)
    n_cached = 0 if cached is None else len(cached)

    if cached is None:
        return fresh_df, f"{market}: fresh {n_fresh} (no prior cache)"
    if n_fresh == 0:
        return cached, f"{market}: fresh scrape EMPTY → kept cached {n_cached} (suspected API issue)"
    if n_fresh < n_cached * 0.5:
        return cached, (f"{market}: fresh {n_fresh} << cached {n_cached} → kept cached "
                        f"(suspected degraded scrape; use refresh to force)")
    return fresh_df, f"{market}: fresh {n_fresh} (replaced cached {n_cached})"


def run_collection(markets=("US", "SGX", "SIX"),
                   pull_financials=True, check_prices=True,
                   use_cache=True, refresh=()):
    """
    use_cache : if True, reuse already-collected events for a market from the
                saved Excel file instead of re-scraping (big time saver — US
                takes ~12-16 min to scrape).
    refresh   : tuple of markets to FORCE re-scrape even if cached
                (e.g. refresh=("SGX",) to re-pull just SGX).
    """
    print("█"*70)
    print("  TAVE EVENT COLLECTION — CODE 1  (with canary + auto-verify + diagnostics)")
    print(f"  {sum(len(FIRMS[m]) for m in markets if m in FIRMS)} firms | {START_DATE} → {END_DATE}")
    print("█"*70)

    # --- 0. Environment: mount Drive (if needed) + create output folder ---
    print("\n── Environment setup ──")
    ensure_drive_and_dirs()

    # --- 1. Collect events (per market, with caching) ---
    market_frames = []     # one DataFrame per market (cached or freshly scraped)
    merge_notes = []        # human-readable per-market provenance for the diagnostics
    for market in markets:
        if market not in FIRMS:
            continue

        # Try cache first (unless this market is in the refresh list)
        if use_cache and market not in refresh:
            cached = _load_cached_market(market)
            if cached is not None and len(cached) > 0:
                # Quality gate: must contain the canary firm to trust the cache
                canary = CANARY_EVENTS.get(market, {})
                firm_ok = True
                if canary:
                    firm_ok = cached["firm_name"].str.contains(
                        canary["firm"], case=False, na=False).any()
                status = "OK" if firm_ok else "canary missing — consider refresh=('%s',)" % market
                print(f"\n── {market} ── (CACHED: {len(cached)} events reused, {status})")
                market_frames.append(cached)
                merge_notes.append(f"{market}: reused cached {len(cached)} ({status})")
                continue

        # Otherwise scrape fresh — collect into a per-market list
        print(f"\n── {market} ── (scraping fresh)")
        collector = COLLECTORS[market]
        fresh_rows = []
        for firm in tqdm(FIRMS[market], desc=market):
            firm = dict(firm); firm["exchange"] = market
            try:
                fresh_rows.extend(collector(firm))
            except Exception as e:
                print(f"    {firm['name']}: {e}")

        fresh_df = pd.DataFrame(fresh_rows)
        if not fresh_df.empty:
            fresh_df["event_date"] = pd.to_datetime(fresh_df["event_date"], errors="coerce")
            fresh_df = fresh_df.dropna(subset=["event_date"])
            # Stamp when this market was collected (provenance for the living doc)
            fresh_df["collected_on"] = datetime.now().strftime("%Y-%m-%d %H:%M")

        # 'Never downgrade' guard: if this fresh scrape is much thinner than the
        # cached version, keep the cache instead (suspected API hiccup).
        chosen, note = _merge_market_keep_best(fresh_df if not fresh_df.empty else None, market)
        merge_notes.append(note)
        if chosen is not None and len(chosen) > 0:
            market_frames.append(chosen)

    # One living dataset: all markets combined
    if market_frames:
        events_df = pd.concat(market_frames, ignore_index=True, sort=False)
    else:
        events_df = pd.DataFrame()

    if merge_notes:
        print("\n  Per-market provenance (living document):")
        for n in merge_notes:
            print(f"    • {n}")

    if not events_df.empty:
        events_df["event_date"] = pd.to_datetime(events_df["event_date"], errors="coerce")
        events_df = events_df.dropna(subset=["event_date"])
        events_df = events_df.sort_values(["firm_name","event_date"]).reset_index(drop=True)

        # Drop any stale id/score columns from cache so we recompute cleanly
        for col in ["event_id", "impact_score", "impact_tier", "dup_group_size",
                    "auto_status", "auto_reason"]:
            if col in events_df.columns:
                events_df = events_df.drop(columns=[col])

        # --- 1a-i. HIGH-IMPACT SCORING (request A) ---
        n_before = len(events_df)
        events_df = score_impact(events_df)

        # --- 1a-ii. DEDUPLICATION (request A) ---
        events_df = deduplicate_events(events_df, window_days=5)
        n_after = len(events_df)
        print(f"\n  Dedup: {n_before} raw filings → {n_after} distinct events "
              f"({n_before - n_after} near-duplicates collapsed)")

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
        # Collection log: per-market provenance + last collection date (living-doc audit trail)
        log_rows = []
        if not events_df.empty and "collected_on" in events_df.columns:
            for mkt in events_df["exchange"].unique():
                sub = events_df[events_df["exchange"] == mkt]
                last = sub["collected_on"].dropna()
                log_rows.append({
                    "market": mkt,
                    "events": len(sub),
                    "last_collected": last.max() if not last.empty else "reused from cache",
                })
        if merge_notes:
            for n in merge_notes:
                log_rows.append({"market": "note", "events": "", "last_collected": n})
        if log_rows:
            pd.DataFrame(log_rows).to_excel(w, sheet_name="Collection_Log", index=False)
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

    # Impact tiers — surfaces the high-impact subset for a cleaner sub-sample
    if "impact_tier" in ver.columns and not ver.empty:
        S.append(f"\nVERIFIED EVENTS BY IMPACT TIER (high-impact = strongest signals)")
        for tier in ["HIGH", "MEDIUM", "LOW"]:
            n = int((ver["impact_tier"] == tier).sum())
            S.append(f"  {tier:<8} {n}")
        hi = int((ver["impact_tier"] == "HIGH").sum())
        S.append(f"  → Consider running the regression on HIGH+MEDIUM impact events for a")
        S.append(f"    cleaner sample ({hi} high-impact events identified).")
        if "dup_group_size" in ver.columns:
            collapsed = int((ver["dup_group_size"] > 1).sum())
            S.append(f"  Events that absorbed near-duplicate filings: {collapsed}")

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


def self_test():
    """
    Fast integrity check — verifies every function the pipeline relies on is
    actually defined, before a long collection run. Catches the 'consumed def
    line' class of bug in <1 second instead of 15 minutes in. Returns True if OK.
    """
    required = [
        "classify_event", "_base_row", "_date_in_range", "_truthy",
        "collect_edgar", "_fetch_edgar_doc_text", "collect_sgx", "collect_six",
        "collect_newsroom", "collect_eqs", "collect_rns", "collect_hkex", "collect_tse",
        "score_impact", "deduplicate_events", "_resolve_cluster",
        "check_canary", "auto_verify", "_load_cached_market", "_merge_market_keep_best",
        "ensure_drive_and_dirs", "get_simfin_financials", "get_yfinance_fundamentals",
        "check_price_coverage", "print_diagnostics", "print_copypaste_summary",
        "run_collection",
    ]
    g = globals()
    missing = [name for name in required if name not in g or not callable(g[name])]
    if missing:
        print("✗ SELF-TEST FAILED — missing/again-broken functions:")
        for m in missing:
            print(f"    - {m}")
        print("  Do NOT run collection until these are restored.")
        return False
    print(f"✓ SELF-TEST PASSED — all {len(required)} pipeline functions defined.")
    return True


if __name__ == "__main__":
    # Recommended first run — US only, to validate EDGAR:
    # run_collection(markets=("US",), pull_financials=False, check_prices=False)

    # Full run:
    run_collection()
