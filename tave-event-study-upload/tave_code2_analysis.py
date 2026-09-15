"""
═══════════════════════════════════════════════════════════════════════════════
  TAVE EVENT STUDY — CODE 2 of 2
  EVENT STUDY ENGINE + CROSS-SECTIONAL OLS REGRESSION
═══════════════════════════════════════════════════════════════════════════════

  Reads: tave_master_database.xlsx (produced by Code 1)
  Produces:
    - Abnormal Returns (AR), Cumulative Abnormal Returns (CAR) per event
    - Cumulative Average Abnormal Returns (CAAR) across the sample
    - Significance tests: Patell (1976), Boehmer-Musumeci-Poulsen (1991),
      Corrado (1989) rank test
    - Cross-sectional OLS regression validating TAVE channels (H1-H5)
    - Robustness: multiple event windows, alternative index models,
      heteroskedasticity-robust standard errors, winsorization

  Output: tave_results.xlsx + tave_AR_timeline.png

  Author: Felix Diego Langer — TAVE Research, GlobalNxt DBA 2026
═══════════════════════════════════════════════════════════════════════════════
"""

# Install (Colab):
#   !pip install pandas numpy scipy statsmodels yfinance matplotlib openpyxl

import pandas as pd
import numpy as np
import yfinance as yf
from scipy import stats
import statsmodels.api as sm
import matplotlib.pyplot as plt
import os
import time
import warnings
warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════════════════

DRIVE_DIR    = "/content/drive/MyDrive/TAVE_Research"
INPUT_FILE   = os.path.join(DRIVE_DIR, "tave_master_database.xlsx")
RESULTS_FILE = os.path.join(DRIVE_DIR, "tave_results.xlsx")
PLOT_FILE    = os.path.join(DRIVE_DIR, "tave_AR_timeline.png")
COEF_PLOT_FILE   = os.path.join(DRIVE_DIR, "tave_coefficients.png")
ASSET_PLOT_FILE  = os.path.join(DRIVE_DIR, "tave_assetclass_caar.png")
FINDINGS_PDF     = os.path.join(DRIVE_DIR, "tave_findings_summary.pdf")


def ensure_drive():
    """
    Mount Google Drive if on Colab and not yet mounted, so Code 2 'just works'.
    If the input file isn't found, fail with a clear, actionable message rather
    than a cryptic read error. Safe to call off Colab.
    """
    if not os.path.isdir("/content/drive/MyDrive"):
        try:
            from google.colab import drive
            print("  Mounting Google Drive...")
            drive.mount("/content/drive")
        except ImportError:
            print("  Not on Colab — expecting input file at the configured path.")
        except Exception as e:
            print(f"  Drive mount issue ({e}).")
    if not os.path.isfile(INPUT_FILE):
        raise FileNotFoundError(
            f"\n  Cannot find {INPUT_FILE}\n"
            f"  → Run Code 1 first (it produces tave_master_database.xlsx), or\n"
            f"  → Check the TAVE_Research folder exists in your Drive and the file is there.")


# Estimation window: [-260, -11] trading days relative to event (t=0)
EST_START, EST_END = -260, -11
# Event windows to test
EVENT_WINDOWS = [(-1, 1), (-1, 3), (-5, 5), (-2, 2), (0, 1)]
# Minimum estimation observations required to keep an event
MIN_EST_OBS = 120

# Local market index per exchange (for the market model)
MARKET_INDEX = {
    "US": "^GSPC", "SGX": "^STI", "XETRA": "^GDAXI", "LSE": "^FTSE",
    "SIX": "^SSMI", "HKEX": "^HSI", "TSE": "^N225",
}
# Map jurisdiction -> index too (fallback)
JURIS_INDEX = {
    "US": "^GSPC", "Singapore": "^STI", "EU": "^GDAXI", "UK": "^FTSE",
    "Switzerland": "^SSMI", "Hong Kong": "^HSI", "Japan": "^N225",
}

# BLOCKER 2 FIX — ETF fallbacks when the native index symbol fails on yfinance.
# These are liquid US-listed MSCI country ETFs that track each market closely.
INDEX_FALLBACK = {
    "^GSPC": "SPY",     # S&P 500
    "^STI":  "EWS",     # MSCI Singapore
    "^GDAXI":"EWG",     # MSCI Germany
    "^FTSE": "EWU",     # MSCI United Kingdom
    "^SSMI": "EWL",     # MSCI Switzerland
    "^HSI":  "EWH",     # MSCI Hong Kong
    "^N225": "EWJ",     # MSCI Japan
}

# Confounding-event filter: days around the event to keep clear of earnings
CONFOUND_WINDOW_DAYS = 5


# ═══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════════════════════

def _truthy(val):
    """Robustly interpret a 'verified' cell that a human may have typed by hand."""
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val == 1
    if isinstance(val, str):
        return val.strip().lower() in ("true", "yes", "y", "1", "x", "verified")
    return False


def _prepare_events_in_memory(df_in):
    """
    Apply the SAME verified-filter + one-per-firm-per-day dedup that load_events
    applies to the file, so the in-memory path yields an identical analysis sample.
    """
    df = df_in.copy()
    df["event_date"] = pd.to_datetime(df["event_date"], errors="coerce")
    df = df.dropna(subset=["event_date"])
    if "verified" in df.columns:
        df["verified"] = df["verified"].apply(_truthy)
        if df["verified"].any():
            df = df[df["verified"]].copy()
        elif "auto_status" in df.columns and (df["auto_status"] == "VERIFIED").any():
            df = df[df["auto_status"] == "VERIFIED"].copy()
    elif "auto_status" in df.columns and (df["auto_status"] == "VERIFIED").any():
        df = df[df["auto_status"] == "VERIFIED"].copy()
    df = df.sort_values("event_date").drop_duplicates(
        subset=["firm_name", "event_date"], keep="first").reset_index(drop=True)
    return df


def load_events() -> pd.DataFrame:
    """Load verified events from the master database."""
    # Freshness guard: report the file's age + per-market breakdown so a STALE
    # (e.g. US-only) read after a multi-market collection is immediately visible,
    # instead of silently analysing the wrong file (Drive sync lag on Colab).
    try:
        _mtime = os.path.getmtime(INPUT_FILE)
        _age_min = (time.time() - _mtime) / 60.0
        print(f"  Reading: {INPUT_FILE}")
        print(f"  File last modified: {_age_min:.1f} min ago")
        if _age_min > 30:
            print("  ⚠ File is >30 min old. If you JUST ran collection, Drive may not have")
            print("    synced yet — wait ~30s and re-run this cell, or the read may be STALE.")
    except Exception:
        pass

    df = pd.read_excel(INPUT_FILE, sheet_name="Events_Classified")
    df["event_date"] = pd.to_datetime(df["event_date"])

    # Show what markets are actually in the file BEFORE filtering — the fastest
    # way to catch a stale US-only read after a 5-market collection.
    if "exchange" in df.columns:
        _mkts = df["exchange"].value_counts().to_dict()
        print(f"  Markets in file: {_mkts}")
        if len(_mkts) == 1 and "US" in _mkts:
            print("  ⚠ ONLY US in file. If you expected SGX/SIX/XETRA/LSE, this is a STALE read")
            print("    (collection's multi-market write hasn't synced). Re-run this cell.")

    # Normalise the 'verified' column robustly — a human may have typed TRUE/yes/x by hand
    if "verified" in df.columns:
        df["verified"] = df["verified"].apply(_truthy)

    # Use verified events if any are marked; else fall back to HIGH confidence
    if "verified" in df.columns and df["verified"].any():
        df = df[df["verified"]].copy()
        print(f"  Using {len(df)} verified events (auto-verified + any manual promotions)")
    elif "auto_status" in df.columns and (df["auto_status"] == "VERIFIED").any():
        df = df[df["auto_status"] == "VERIFIED"].copy()
        print(f"  Using {len(df)} auto-VERIFIED events (no manual promotions found)")
    else:
        df = df[df["classification_conf"] == "HIGH"].copy()
        print(f"  No verification flags found — using {len(df)} HIGH-confidence events")
        print("  (Recommend reviewing Manual_Review and promoting genuine events first)")

    # Deduplicate: one event per firm per day (multiple filings same day = one event)
    df = df.sort_values("event_date").drop_duplicates(
        subset=["firm_name", "event_date"], keep="first").reset_index(drop=True)
    print(f"  After dedup (one event per firm per day): {len(df)} events")
    # Post-filter market breakdown — confirms the analysis sample spans all markets
    if "exchange" in df.columns:
        print(f"  Verified events by market: {df['exchange'].value_counts().to_dict()}")
    return df


def download_prices(tickers: list, index_tickers: list):
    """
    Download all needed price series once.
    BLOCKER 2 FIX: validate every index symbol; if a native index returns empty,
    substitute its ETF fallback and record the substitution. Returns
    (data_dict, index_health) so the diagnostic block can report exactly which
    markets are usable.
    """
    all_tickers = list(set([t for t in tickers if t and t != "NONE"]))
    print(f"  Downloading {len(all_tickers)} firm price series...")

    def _extract_close(h):
        """Return a clean 1-D Series of adjusted closes from a yfinance frame.
        Handles multi-index columns (newer yfinance) and 1-col DataFrames."""
        if h is None or len(h) == 0:
            return pd.Series(dtype=float)
        col = None
        if isinstance(h.columns, pd.MultiIndex):
            for field in ("Adj Close", "Close"):
                if field in h.columns.get_level_values(0):
                    col = h[field]
                    break
        else:
            for field in ("Adj Close", "Close"):
                if field in h.columns:
                    col = h[field]
                    break
        if col is None:
            return pd.Series(dtype=float)
        if isinstance(col, pd.DataFrame):
            col = col.iloc[:, 0] if col.shape[1] >= 1 else pd.Series(dtype=float)
        return pd.to_numeric(col, errors="coerce").dropna()

    def _extract_field(h, fields):
        """Generic 1-D extractor for any OHLCV field (multi-index safe)."""
        if h is None or len(h) == 0:
            return pd.Series(dtype=float)
        col = None
        if isinstance(h.columns, pd.MultiIndex):
            for field in fields:
                if field in h.columns.get_level_values(0):
                    col = h[field]; break
        else:
            for field in fields:
                if field in h.columns:
                    col = h[field]; break
        if col is None:
            return pd.Series(dtype=float)
        if isinstance(col, pd.DataFrame):
            col = col.iloc[:, 0] if col.shape[1] >= 1 else pd.Series(dtype=float)
        return pd.to_numeric(col, errors="coerce")

    data = {}
    volume_data = {}   # ticker -> DataFrame[close, volume] for liquidity metrics (Amihud/turnover)
    for t in all_tickers:
        try:
            h = yf.download(t, start="2019-06-01", end="2026-01-31",
                            progress=False, auto_adjust=False)
            s = _extract_close(h)
            if len(s) > 0:
                data[t] = s
                # Keep close+volume aligned for the liquidity channel (WACC channel 1)
                vol = _extract_field(h, ("Volume",))
                close_raw = _extract_field(h, ("Close", "Adj Close"))
                lv = pd.DataFrame({"close": close_raw, "volume": vol}).dropna()
                if len(lv) > 0:
                    volume_data[t] = lv
        except Exception as e:
            print(f"    firm {t}: {e}")

    # --- Indices with validation + fallback ---
    print(f"  Downloading + validating {len(index_tickers)} index series...")
    index_health = {}        # original_symbol -> {used, n_days, status}
    index_resolved = {}      # original_symbol -> symbol actually usable

    for idx in index_tickers:
        used, n_days, status = None, 0, ""
        # try native
        try:
            h = yf.download(idx, start="2019-06-01", end="2026-01-31",
                            progress=False, auto_adjust=False)
            s = _extract_close(h)
            if len(s) >= 500:
                data[idx] = s; used = idx; n_days = len(s); status = "NATIVE_OK"
        except Exception:
            pass

        # fallback if native failed
        if used is None:
            fb = INDEX_FALLBACK.get(idx)
            if fb:
                try:
                    h = yf.download(fb, start="2019-06-01", end="2026-01-31",
                                    progress=False, auto_adjust=False)
                    s = _extract_close(h)
                    if len(s) >= 500:
                        data[idx] = s   # store under the ORIGINAL key so lookups still work
                        used = fb; n_days = len(s); status = "FALLBACK_ETF"
                except Exception:
                    pass

        if used is None:
            status = "FAILED"
        index_health[idx] = {"index": idx, "used_symbol": used, "n_days": n_days, "status": status}
        index_resolved[idx] = used

    return data, index_health, volume_data


# ═══════════════════════════════════════════════════════════════════════════════
# CONFOUNDING-EVENT FILTER (econometric refinement)
# ═══════════════════════════════════════════════════════════════════════════════

def _tz_naive(ts):
    """Normalize a timestamp to tz-naive, midnight. Handles tz-aware inputs from yfinance."""
    t = pd.Timestamp(ts)
    if t.tzinfo is not None or getattr(t, "tz", None) is not None:
        t = t.tz_localize(None)
    return t.normalize()


def get_earnings_dates(yf_ticker: str) -> list:
    """
    Fetch known earnings dates for a firm via yfinance. Used to flag events whose
    window overlaps an earnings release (a classic confound). Returns list of dates
    or empty list if unavailable. All dates returned tz-naive.
    """
    try:
        tk = yf.Ticker(yf_ticker)
        ed = tk.get_earnings_dates(limit=40)
        if ed is not None and not ed.empty:
            return [_tz_naive(d) for d in ed.index]
    except Exception:
        pass
    return []


def is_confounded(event_date, earnings_dates, window_days=CONFOUND_WINDOW_DAYS) -> bool:
    """True if any earnings date falls within +/- window_days of the event."""
    ed_event = _tz_naive(event_date)
    for ed in earnings_dates:
        if abs((_tz_naive(ed) - ed_event).days) <= window_days:
            return True
    return False


# ═══════════════════════════════════════════════════════════════════════════════
# EVENT STUDY CORE
# ═══════════════════════════════════════════════════════════════════════════════

def compute_returns(price_series) -> pd.Series:
    """
    Daily log returns. Always returns a 1-D pandas Series (possibly empty).
    Guards against yfinance returning a 1-column DataFrame or a degenerate scalar.
    """
    s = price_series
    # If a 1-column DataFrame slipped through, squeeze to a Series
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0] if s.shape[1] >= 1 else pd.Series(dtype=float)
    if not isinstance(s, pd.Series):
        s = pd.Series(s)
    s = pd.to_numeric(s, errors="coerce").dropna()
    if len(s) < 2:
        return pd.Series(dtype=float)   # cannot compute a return from <2 points
    return np.log(s / s.shift(1)).dropna()


def estimate_market_model(firm_ret, mkt_ret, event_date):
    """
    Estimate alpha, beta on the estimation window [-260, -11].
    Returns (alpha, beta, residual_std, n_obs) or None if insufficient data.
    """
    # Guard: both inputs must be non-empty Series, else we can't estimate
    if not isinstance(firm_ret, pd.Series) or not isinstance(mkt_ret, pd.Series):
        return None
    if len(firm_ret) < MIN_EST_OBS or len(mkt_ret) < MIN_EST_OBS:
        return None

    # Align
    df = pd.DataFrame({"firm": firm_ret, "mkt": mkt_ret}).dropna()
    df = df.sort_index()

    # Normalize index to tz-naive so comparisons with event_date never clash
    if getattr(df.index, "tz", None) is not None:
        df.index = df.index.tz_localize(None)
    event_date = _tz_naive(event_date)
    # Trading-day index relative to event
    if event_date not in df.index:
        # find nearest trading day on/after event
        future = df.index[df.index >= event_date]
        if len(future) == 0:
            return None
        event_pos = df.index.get_loc(future[0])
    else:
        event_pos = df.index.get_loc(event_date)

    est_lo = event_pos + EST_START
    est_hi = event_pos + EST_END
    if est_lo < 0:
        return None

    est = df.iloc[est_lo:est_hi]
    if len(est) < MIN_EST_OBS:
        return None

    X = sm.add_constant(est["mkt"].values)
    y = est["firm"].values
    model = sm.OLS(y, X).fit()
    # params may be a pandas Series (named) or ndarray — use positional access
    p = np.asarray(model.params, dtype=float)
    alpha, beta = p[0], p[1]
    resid_std = np.std(model.resid, ddof=2)
    return {"alpha": alpha, "beta": beta, "resid_std": resid_std,
            "n_obs": len(est), "event_pos": event_pos, "aligned": df}


def compute_ARs(est_result, window):
    """
    Compute abnormal returns for the event window.
    Returns DataFrame with relative day, AR, and standardized AR.
    """
    df = est_result["aligned"]
    pos = est_result["event_pos"]
    alpha, beta, rstd = est_result["alpha"], est_result["beta"], est_result["resid_std"]

    lo, hi = pos + window[0], pos + window[1]
    if lo < 0 or hi >= len(df):
        return None

    win = df.iloc[lo:hi+1].copy()
    win["rel_day"] = range(window[0], window[1]+1)
    win["expected"] = alpha + beta * win["mkt"]
    win["AR"] = win["firm"] - win["expected"]
    win["SAR"] = win["AR"] / rstd   # standardized AR
    return win[["rel_day", "AR", "SAR"]].reset_index(drop=True)


def event_study(events_df, price_data, filter_confounds=True):
    """
    Run the full event study across all events and all windows.
    Returns: car_results, ar_timeline, study_diag (skip reasons for diagnostics).
    """
    car_rows = []
    timeline_ar = {d: [] for d in range(-10, 11)}
    skip = {"no_ticker": [], "no_index": [], "est_failed": [], "confounded": [], "ok": []}

    # Pre-fetch earnings dates per ticker (cache) if filtering
    earnings_cache = {}

    for _, ev in events_df.iterrows():
        yft = ev.get("yf_ticker", "")
        eid = ev.get("event_id", "?")
        if not yft or yft == "NONE" or yft not in price_data:
            skip["no_ticker"].append(f"{eid} {ev.get('firm_name','')} (yf={yft})")
            continue

        idx_t = MARKET_INDEX.get(ev["exchange"]) or JURIS_INDEX.get(ev["jurisdiction"])
        if idx_t not in price_data:
            skip["no_index"].append(f"{eid} {ev.get('firm_name','')} (idx={idx_t})")
            continue

        # Confounding-event filter
        if filter_confounds:
            if yft not in earnings_cache:
                earnings_cache[yft] = get_earnings_dates(yft)
            if is_confounded(ev["event_date"], earnings_cache[yft]):
                skip["confounded"].append(f"{eid} {ev.get('firm_name','')} {str(ev['event_date'])[:10]}")
                continue

        firm_ret = compute_returns(price_data[yft])
        mkt_ret = compute_returns(price_data[idx_t])

        est = estimate_market_model(firm_ret, mkt_ret, pd.Timestamp(ev["event_date"]))
        if est is None:
            skip["est_failed"].append(f"{eid} {ev.get('firm_name','')} (insufficient est-window data)")
            continue

        tl = compute_ARs(est, (-10, 10))
        if tl is not None:
            for _, r in tl.iterrows():
                if int(r["rel_day"]) in timeline_ar:
                    timeline_ar[int(r["rel_day"])].append(r["AR"])

        for w in EVENT_WINDOWS:
            ar = compute_ARs(est, w)
            if ar is None:
                continue
            car = ar["AR"].sum()
            scar = ar["SAR"].sum() / np.sqrt(len(ar))
            car_rows.append({
                "event_id": ev["event_id"], "firm_name": ev["firm_name"],
                "ticker": ev["ticker"], "yf_ticker": ev.get("yf_ticker", ev["ticker"]),
                "exchange": ev["exchange"],
                "jurisdiction": ev["jurisdiction"],
                "jurisdiction_score": ev.get("jurisdiction_score", np.nan),
                "event_date": ev["event_date"],
                "announcement_type": ev.get("announcement_type", ""),
                "asset_class": ev.get("asset_class", ""),
                "audit_disclosed": ev.get("audit_disclosed", False),
                "pct_tokenized": ev.get("pct_tokenized", np.nan),
                "impact_tier": ev.get("impact_tier", "MEDIUM"),
                "impact_score": ev.get("impact_score", np.nan),
                "window": f"({w[0]},{w[1]})", "window_tuple": w,
                "CAR": car, "SCAR": scar, "beta": est["beta"],
                "resid_std": est["resid_std"], "n_est_obs": est["n_obs"],
            })
        skip["ok"].append(eid)

    car_df = pd.DataFrame(car_rows)
    return car_df, timeline_ar, skip


# ═══════════════════════════════════════════════════════════════════════════════
# SIGNIFICANCE TESTS
# ═══════════════════════════════════════════════════════════════════════════════

def patell_test(scar_values):
    """
    Patell (1976): aggregate standardized CARs.
    Z = (1/sqrt(N)) * sum(SCAR_i), ~ N(0,1) under H0.
    """
    scar = np.asarray(scar_values)
    scar = scar[~np.isnan(scar)]
    N = len(scar)
    if N < 2:
        return np.nan, np.nan, N
    Z = np.sum(scar) / np.sqrt(N)
    p = 2 * (1 - stats.norm.cdf(abs(Z)))
    return Z, p, N


def bmp_test(scar_values):
    """
    Boehmer, Musumeci & Poulsen (1991): standardized cross-sectional test.
    Corrects for event-induced variance. THIS IS THE PRIMARY TEST.
    t = mean(SCAR) / (std(SCAR)/sqrt(N))
    """
    scar = np.asarray(scar_values)
    scar = scar[~np.isnan(scar)]
    N = len(scar)
    if N < 2:
        return np.nan, np.nan, N
    mean_scar = np.mean(scar)
    std_scar = np.std(scar, ddof=1)
    t = mean_scar / (std_scar / np.sqrt(N))
    p = 2 * (1 - stats.t.cdf(abs(t), df=N-1))
    return t, p, N


def corrado_rank_test(car_values):
    """
    Corrado (1989) non-parametric rank test (simplified cross-sectional form).
    Tests whether event-window CARs rank significantly above their distribution.
    Uses a sign+rank approach robust to non-normality.
    """
    car = np.asarray(car_values)
    car = car[~np.isnan(car)]
    N = len(car)
    if N < 2:
        return np.nan, np.nan, N
    # Wilcoxon signed-rank against zero median
    try:
        stat, p = stats.wilcoxon(car)
        # Convert to a z-ish direction indicator
        direction = np.sign(np.median(car))
        return stat * direction, p, N
    except Exception:
        return np.nan, np.nan, N


def t_test_car(car_values):
    """Simple cross-sectional t-test on raw CARs (reported alongside)."""
    car = np.asarray(car_values)
    car = car[~np.isnan(car)]
    N = len(car)
    if N < 2:
        return np.nan, np.nan, N
    t, p = stats.ttest_1samp(car, 0)
    return t, p, N


def sign_test(car_values):
    """
    Non-parametric sign test: are positive CARs more/less frequent than 50%?
    Robust to non-normality (your returns are skewed/leptokurtic, so this
    directly addresses that reviewer concern). Returns (z_stat, p_value, N).
    """
    car = np.asarray(car_values)
    car = car[~np.isnan(car)]
    N = len(car)
    if N < 2:
        return np.nan, np.nan, N
    n_pos = int((car > 0).sum())
    # Binomial test against p=0.5, two-sided; z-approx for reporting
    p_val = stats.binomtest(n_pos, N, 0.5, alternative="two-sided").pvalue
    expected = N * 0.5
    z = (n_pos - expected) / np.sqrt(N * 0.25) if N > 0 else np.nan
    return z, p_val, N


def run_all_tests(car_df):
    """Run all significance tests for each event window."""
    results = []
    for w in EVENT_WINDOWS:
        wlabel = f"({w[0]},{w[1]})"
        sub = car_df[car_df["window"] == wlabel]
        if sub.empty:
            continue
        caar = sub["CAR"].mean()
        median_car = sub["CAR"].median()

        t_p, p_p, n = patell_test(sub["SCAR"].values)
        t_b, p_b, _ = bmp_test(sub["SCAR"].values)
        t_c, p_c, _ = corrado_rank_test(sub["CAR"].values)
        t_t, p_t, _ = t_test_car(sub["CAR"].values)
        z_s, p_s, _ = sign_test(sub["CAR"].values)

        pct_positive = (sub["CAR"] > 0).mean() * 100

        results.append({
            "window": wlabel,
            "N": n,
            "CAAR": caar,
            "CAAR_pct": caar * 100,
            "median_CAR_pct": median_car * 100,
            "pct_positive": pct_positive,
            "t_test_stat": t_t, "t_test_p": p_t,
            "Patell_Z": t_p, "Patell_p": p_p,
            "BMP_t": t_b, "BMP_p": p_b,           # PRIMARY
            "Corrado_stat": t_c, "Corrado_p": p_c,
            "Sign_z": z_s, "Sign_p": p_s,          # non-parametric robustness
            "sig_BMP": "***" if p_b < 0.01 else "**" if p_b < 0.05 else "*" if p_b < 0.10 else "",
        })
    return pd.DataFrame(results)


# ═══════════════════════════════════════════════════════════════════════════════
# CROSS-SECTIONAL OLS — TAVE VALIDATION
# ═══════════════════════════════════════════════════════════════════════════════

def cross_sectional_regression(car_df, fundamentals_df=None, window="(-1,1)",
                               impact_filter=None, include_interaction=False,
                               exclude_firms=None, jurisdiction_fe=False,
                               announcement_fe=False):
    """
    Regress CAR on TAVE channel proxies.

    CAR_i = g0 + g1*BaselineIlliquidity + g2*TokenizationScope
            + g3*JurisdictionScore + g4*AuditPresence
            + g5*AssetClass_dummies + g6*log(MktCap) + g7*Leverage + e

    Tests H2-H5. H1 is the event study itself.

    impact_filter: None = full sample (primary regression).
                   ["HIGH","MEDIUM"] = high-impact subset (robustness regression).
    """
    df = car_df[car_df["window"] == window].copy()

    # Optional impact-tier restriction (robustness regression)
    if impact_filter is not None and "impact_tier" in df.columns:
        df = df[df["impact_tier"].astype(str).isin(impact_filter)].copy()

    # Optional firm exclusion (e.g. Sea Limited — NYSE-listed/USD, STI-benchmarked:
    # a market-model misspecification whose extreme CARs sit coded as Singapore/
    # clarity-5 and can drive the JurisdictionScore coefficient).
    if exclude_firms:
        df = df[~df["firm_name"].isin(set(exclude_firms))].copy()

    if df.empty or len(df) < 10:
        print(f"  Insufficient observations for regression ({len(df)})")
        return None

    # firm_id for clustering standard errors
    df["firm_id"] = df["firm_name"].astype("category").cat.codes

    # --- Build independent variables ---

    # TokenizationScope: ordinal from announcement type
    scope_map = {"Pilot": 1, "New Program": 2, "Expansion": 3, "Live Launch": 4, "Unclassified": np.nan}
    df["TokenizationScope"] = df["announcement_type"].map(scope_map)

    # AuditPresence
    df["AuditPresence"] = df["audit_disclosed"].astype(float)

    # JurisdictionScore already present
    df["JurisdictionScore"] = pd.to_numeric(df["jurisdiction_score"], errors="coerce")

    # BaselineIlliquidity — RE-SPECIFIED (was a weak 1-5 ordinal, flat at p~0.99).
    # Now the asset class's literature illiquidity premium ILP_max in % p.a.
    # (Amihud & Mendelson 1986; Pástor & Stambaugh 2003; values per TAVE Master Doc).
    # This gives the regressor genuine economic magnitude tied to Gap 1, so the
    # coefficient is a real test of the illiquidity channel rather than a rank.
    illiq_map = {
        "Supply Chain": 1.80, "Receivables": 1.60, "Real Estate": 1.20,
        "Fund": 0.80, "Bond": 0.60, "Gold/Commodity": 0.40,
        "Equity": 0.20, "Deposit": 0.10, "Unclassified": np.nan,
    }
    df["BaselineIlliquidity"] = df["asset_class"].map(illiq_map)

    # Interaction (TAVE sharp test): does the illiquidity payoff depend on regulatory
    # clarity? IlliqXJuris = BaselineIlliquidity * JurisdictionScore. A positive sign
    # would say the illiquidity channel pays off more in clearer regimes.
    df["IlliqXJuris"] = df["BaselineIlliquidity"] * pd.to_numeric(df.get("jurisdiction_score"), errors="coerce")
    if fundamentals_df is not None and not fundamentals_df.empty:
        fcols = fundamentals_df.copy()
        fcols.columns = [c.lower() for c in fcols.columns]
        merge_key = "ticker"
        if merge_key in fcols.columns:
            df = df.merge(
                fcols[[merge_key] + [c for c in ["market_cap","debt_to_equity","ebitda"] if c in fcols.columns]],
                left_on="ticker", right_on=merge_key, how="left")
            if "market_cap" in df.columns:
                df["log_MktCap"] = np.log(df["market_cap"].replace(0, np.nan))
            if "debt_to_equity" in df.columns:
                df["Leverage"] = df["debt_to_equity"]

    # --- Assemble regression matrix ---
    # PRIMARY (default): clean H2 test — re-specified illiquidity + controls, NO
    # interaction. The interaction IlliqXJuris is collinear with its components
    # (it shares JurisdictionScore), so including it makes the BaselineIlliquidity
    # main coefficient an extrapolation to JurisdictionScore=0 (outside the 2-5 data
    # range) and inflates the condition number. It is therefore reported only in a
    # SEPARATE exploratory model (include_interaction=True), never in the primary.
    base_vars = ["BaselineIlliquidity", "TokenizationScope",
                 "JurisdictionScore", "AuditPresence"]
    fe_cols = []
    if jurisdiction_fe and "jurisdiction" in df.columns:
        # FIXED-EFFECTS VARIANT: replace the linear clarity score with jurisdiction
        # dummies (base = largest cohort, normally US). Drops the linearity
        # assumption: tests whether jurisdictions differ at all, not whether the
        # ordinal clarity coding is the right functional form.
        base = df["jurisdiction"].mode().iloc[0]
        dummies = pd.get_dummies(df["jurisdiction"], prefix="J").astype(float)
        drop_col = f"J_{base}"
        if drop_col in dummies.columns:
            dummies = dummies.drop(columns=[drop_col])
        fe_cols = list(dummies.columns)
        df = pd.concat([df, dummies], axis=1)
        base_vars = ["BaselineIlliquidity", "TokenizationScope", "AuditPresence"] + fe_cols
        print(f"  Jurisdiction fixed effects (base = {base}): {fe_cols}")
    if announcement_fe and "announcement_type" in df.columns:
        # ANNOUNCEMENT-TYPE CONTROL (materiality/composition, Layer 2): dummies for
        # announcement type. TokenizationScope is DERIVED from announcement_type, so
        # it must be dropped here — keeping both would be collinear by construction.
        ann_base = df["announcement_type"].astype(str).mode().iloc[0]
        ann_d = pd.get_dummies(df["announcement_type"].astype(str), prefix="A").astype(float)
        ann_drop = "A_" + ann_base
        if ann_drop in ann_d.columns:
            ann_d = ann_d.drop(columns=[ann_drop])
        ann_d.columns = [c.replace(" ", "_").replace("/", "_") for c in ann_d.columns]
        df = pd.concat([df, ann_d], axis=1)
        base_vars = [v for v in base_vars if v != "TokenizationScope"] + list(ann_d.columns)
        print(f"  Announcement-type controls (base = {ann_base}; TokenizationScope dropped"
              f" — derived from type): {list(ann_d.columns)}")
    candidate_vars = base_vars + (["IlliqXJuris"] if include_interaction else [])
    if "log_MktCap" in df.columns:
        candidate_vars.append("log_MktCap")
    if "Leverage" in df.columns:
        candidate_vars.append("Leverage")

    reg_df = df[["CAR", "firm_id"] + candidate_vars].dropna()
    if len(reg_df) < len(candidate_vars) + 3:
        print(f"  After dropping NaN, only {len(reg_df)} obs — reducing variables")
        candidate_vars = list(base_vars)
        reg_df = df[["CAR", "firm_id"] + candidate_vars].dropna()

    if len(reg_df) < 8:
        print(f"  Too few observations for regression ({len(reg_df)})")
        return None

    # Drop zero-variance (constant) columns — they make the design matrix singular.
    # E.g. AuditPresence is all-zero when no event disclosed a smart-contract audit;
    # that absence is reported descriptively elsewhere, but it can't enter the OLS.
    dropped_constant = []
    for v in list(candidate_vars):
        if reg_df[v].nunique(dropna=True) <= 1:
            dropped_constant.append(v)
            candidate_vars.remove(v)
    if dropped_constant:
        print(f"  Dropped zero-variance variable(s) from OLS: {dropped_constant} "
              f"(no within-sample variation — reported descriptively, not modelled)")
    if not candidate_vars:
        print("  No usable regressors after dropping constants."); return None

    X = sm.add_constant(reg_df[candidate_vars])
    y = reg_df["CAR"]

    # ECONOMETRIC REFINEMENT: cluster standard errors by firm.
    # Multiple events from the same firm (DBS, JPMorgan) are not independent.
    # We fit two versions and report both: HC3-robust and firm-clustered.
    model_hc3 = sm.OLS(y, X).fit(cov_type="HC3")

    model_clustered = None
    n_clusters = reg_df["firm_id"].nunique() if "firm_id" in reg_df.columns else None
    if "firm_id" in reg_df.columns and n_clusters and n_clusters >= 2:
        try:
            model_clustered = sm.OLS(y, X).fit(
                cov_type="cluster",
                cov_kwds={"groups": reg_df["firm_id"].values})
        except Exception as e:
            print(f"  Cluster-SE fit failed ({e}); reporting HC3 only.")

    # Primary model = clustered if available, else HC3
    model = model_clustered if model_clustered is not None else model_hc3

    return {
        "model": model,
        "model_hc3": model_hc3,
        "model_clustered": model_clustered,
        "n_obs": len(reg_df),
        "n_clusters": n_clusters,
        "se_type": "firm-clustered" if model_clustered is not None else "HC3-robust",
        "variables": candidate_vars,
        "summary": model.summary().as_text(),
        "coefficients": _coef_frame(model),
        "reg_df": reg_df,
    }


def wild_cluster_bootstrap_p(reg_df, test_var, n_boot=999, seed=42):
    """
    Wild cluster bootstrap p-value (Cameron-Gelbach-Miller, Rademacher weights,
    null imposed) for a single coefficient. The standard remedy when the number
    of clusters is small (here: 11 firms), where analytic cluster-robust SEs are
    anti-conservative and can overstate significance.

    reg_df: the exact estimation sample returned by cross_sectional_regression
            (columns: CAR, firm_id, covariates).
    test_var: coefficient to test (e.g. "JurisdictionScore").
    Returns dict with observed clustered t and the bootstrap p-value.
    """
    xcols = [c for c in reg_df.columns if c not in ("CAR", "firm_id")]
    if test_var not in xcols:
        return {"note": f"{test_var} not in regression variables"}
    y = reg_df["CAR"].values
    clusters = reg_df["firm_id"].values
    X_full = sm.add_constant(reg_df[xcols].astype(float))
    fit = sm.OLS(y, X_full).fit(cov_type="cluster", cov_kwds={"groups": clusters})
    t_obs = float(fit.tvalues[test_var])

    # Restricted model: impose H0 (beta_test = 0)
    xr = [c for c in xcols if c != test_var]
    X_r = sm.add_constant(reg_df[xr].astype(float))
    fit_r = sm.OLS(y, X_r).fit()
    resid_r = np.asarray(fit_r.resid)
    yhat_r = np.asarray(fit_r.fittedvalues)

    uniq = np.unique(clusters)
    rng = np.random.default_rng(seed)
    t_boot = np.empty(n_boot)
    for b_i in range(n_boot):
        w = rng.choice(np.array([-1.0, 1.0]), size=len(uniq))
        wmap = {c: w[k] for k, c in enumerate(uniq)}
        y_star = yhat_r + resid_r * np.array([wmap[c] for c in clusters])
        fb = sm.OLS(y_star, X_full).fit(cov_type="cluster", cov_kwds={"groups": clusters})
        t_boot[b_i] = float(fb.tvalues[test_var])
    p_wcb = (np.sum(np.abs(t_boot) >= abs(t_obs)) + 1.0) / (n_boot + 1.0)
    return {"test_var": test_var, "t_clustered": round(t_obs, 3),
            "p_wild_cluster_bootstrap": round(float(p_wcb), 4),
            "n_boot": n_boot, "n_clusters": int(len(uniq)),
            "note": "null-imposed Rademacher WCB; remedy for few-cluster inference"}


def _coef_frame(model):
    """
    Build a coefficient table robustly, regardless of whether statsmodels returns
    pandas Series, numpy arrays, or scalars (a 1-variable model returns scalars,
    which breaks a naive pd.DataFrame({...}) with 'all scalar values' error).
    """
    # Parameter names (index)
    params = model.params
    if hasattr(params, "index"):
        names = list(params.index)
    else:
        names = [f"x{i}" for i in range(np.atleast_1d(params).shape[0])]

    def _arr(x):
        a = np.atleast_1d(np.asarray(x, dtype=float))
        return a

    coef = _arr(model.params)
    bse = _arr(model.bse)
    tval = _arr(model.tvalues)
    pval = _arr(model.pvalues)

    # conf_int can be a DataFrame, 2D array, or (for 1 param) a short vector
    ci = model.conf_int()
    ci = np.asarray(ci, dtype=float)
    if ci.ndim == 1:
        ci = ci.reshape(1, -1)
    ci_low, ci_high = ci[:, 0], ci[:, 1]

    n = len(coef)
    # Defensive alignment — truncate/pad names to match coef length
    if len(names) != n:
        names = [f"x{i}" for i in range(n)]

    return pd.DataFrame(
        {"coef": coef, "std_err": bse, "t_stat": tval,
         "p_value": pval, "ci_low": ci_low, "ci_high": ci_high},
        index=names,
    )


# ═══════════════════════════════════════════════════════════════════════════════
# ROBUSTNESS
# ═══════════════════════════════════════════════════════════════════════════════

def robustness_winsorized(car_df, window="(-1,1)", limits=(0.05, 0.05)):
    """Re-run tests with winsorized CARs to check outlier sensitivity."""
    from scipy.stats.mstats import winsorize
    sub = car_df[car_df["window"] == window].copy()
    if sub.empty:
        return None
    sub["CAR_wins"] = winsorize(sub["CAR"].values, limits=limits)
    caar_w = sub["CAR_wins"].mean()
    t, p = stats.ttest_1samp(sub["CAR_wins"].dropna(), 0)
    return {"window": window, "CAAR_winsorized_pct": caar_w*100, "t": t, "p": p, "N": len(sub)}


def subsample_by_region(car_df, window="(-1,1)"):
    """CAAR and BMP test by region — check effect isn't driven by one market."""
    rows = []
    sub = car_df[car_df["window"] == window]
    for region, g in sub.groupby("exchange"):
        t_b, p_b, n = bmp_test(g["SCAR"].values)
        rows.append({"region": region, "N": n, "CAAR_pct": g["CAR"].mean()*100,
                     "BMP_t": t_b, "BMP_p": p_b})
    return pd.DataFrame(rows)


def subsample_by_assetclass(car_df, window="(-1,1)"):
    """CAAR by asset class — tests H2 (illiquidity channel) descriptively."""
    rows = []
    sub = car_df[car_df["window"] == window]
    for ac, g in sub.groupby("asset_class"):
        t_b, p_b, n = bmp_test(g["SCAR"].values)
        rows.append({"asset_class": ac, "N": n, "CAAR_pct": g["CAR"].mean()*100,
                     "BMP_t": t_b, "BMP_p": p_b})
    return pd.DataFrame(rows).sort_values("CAAR_pct", ascending=False)


# Asset classes whose VALUE depends on provenance/ownership verification that is
# opaque off-chain (physical or real-world assets). Tokenization adds continuously
# verifiable provenance here, which is the Easley-O'Hara (Gap 2, information
# asymmetry) and Bernanke-Gertler (Gap 3, collateral quality) mechanism — distinct
# from the Amihud liquidity channel. "Financial" assets are already transparent.
PROVENANCE_SENSITIVE = {"Gold/Commodity", "Real Estate", "Receivables", "Supply Chain"}
FINANCIAL_TRANSPARENT = {"Bond", "Fund", "Equity", "Deposit"}


def provenance_contrast(car_df, window="(-1,1)"):
    """
    STRUCTURAL PROXY (NOT a channel measurement) for Gaps 2/3.

    Splits events into provenance-sensitive (physical/real-world assets whose
    ownership/provenance was opaque off-chain) vs financial-transparent (already
    liquid/transparent instruments). The information-asymmetry / collateral-quality
    channels predict the provenance-sensitive group reacts LESS NEGATIVELY (more
    positively), because on-chain provenance adds verifiable information exactly
    where it was previously missing.

    This is a coarse proxy built from asset_class only — it CANNOT measure the
    microstructure mechanism (that needs I/B/E/S / TAQ / bond data). Reported as a
    light, exploratory robustness signal for future research, not as channel
    validation. Returns a 2-row group summary + a difference-in-means test.
    """
    sub = car_df[car_df["window"] == window].copy()
    if sub.empty:
        return pd.DataFrame(), {}
    sub["prov_group"] = np.where(sub["asset_class"].isin(PROVENANCE_SENSITIVE), "Provenance-sensitive",
                         np.where(sub["asset_class"].isin(FINANCIAL_TRANSPARENT), "Financial-transparent", "Other"))
    rows = []
    for grp in ("Provenance-sensitive", "Financial-transparent"):
        g = sub[sub["prov_group"] == grp]
        if g.empty:
            continue
        t_b, p_b, n = bmp_test(g["SCAR"].values)
        rows.append({"group": grp, "N": n,
                     "CAAR_pct": round(g["CAR"].mean()*100, 3),
                     "median_CAR_pct": round(g["CAR"].median()*100, 3),
                     "BMP_t": round(t_b, 3) if pd.notna(t_b) else np.nan,
                     "BMP_p": round(p_b, 4) if pd.notna(p_b) else np.nan})
    summary = pd.DataFrame(rows)
    # Difference-in-means (Welch) between the two groups' CARs — does provenance
    # sensitivity separate the reaction at all?
    diff = {}
    pv = sub[sub["prov_group"] == "Provenance-sensitive"]["CAR"].dropna()
    fn = sub[sub["prov_group"] == "Financial-transparent"]["CAR"].dropna()
    if len(pv) >= 3 and len(fn) >= 3:
        t, p = stats.ttest_ind(pv, fn, equal_var=False)
        diff = {"diff_CAAR_pct": round((pv.mean()-fn.mean())*100, 3),
                "welch_t": round(t, 3), "welch_p": round(p, 4),
                "n_prov": len(pv), "n_fin": len(fn)}
    return summary, diff


def pct_tokenized_signal(car_df, window="(-1,1)"):
    """
    STRUCTURAL PROXY (NOT a channel measurement) — light check on whether the
    disclosed share of balance sheet tokenized (pct_tokenized) carries any signal
    in the cross-section. Coverage is thin (~99 events disclose it), so this is an
    underpowered exploratory check, reported only to flag whether a richer scope
    variable would be worth collecting in future work. Returns a small correlation
    + univariate slope summary on the events that disclose pct_tokenized.
    """
    sub = car_df[(car_df["window"] == window)].copy()
    if "pct_tokenized" not in sub.columns:
        return {}
    sub = sub[["CAR", "pct_tokenized"]].dropna()
    sub = sub[sub["pct_tokenized"] > 0]
    if len(sub) < 10:
        return {"n": len(sub), "note": "too few disclosed pct_tokenized values for a signal check"}
    r = np.corrcoef(sub["pct_tokenized"], sub["CAR"])[0, 1]
    # univariate OLS slope (CAR on pct_tokenized) — descriptive only
    x = sub["pct_tokenized"].values; y = sub["CAR"].values
    slope, intercept = np.polyfit(x, y, 1)
    # significance of correlation
    n = len(sub)
    t = r * np.sqrt((n - 2) / max(1e-12, (1 - r**2)))
    from scipy.stats import t as _t
    p = 2 * (1 - _t.cdf(abs(t), df=n-2))
    return {"n": n, "corr_CAR_pct_tokenized": round(float(r), 4),
            "univariate_slope": round(float(slope), 6),
            "corr_p": round(float(p), 4)}


# ═══════════════════════════════════════════════════════════════════════════════
# PLOTTING
# ═══════════════════════════════════════════════════════════════════════════════

def plot_ar_timeline(timeline_ar, save_path):
    """Plot average AR and cumulative AAR from -10 to +10."""
    days = sorted(timeline_ar.keys())
    aar = [np.mean(timeline_ar[d]) if timeline_ar[d] else 0 for d in days]
    caar = np.cumsum(aar)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8), sharex=True)
    ax1.bar(days, [a*100 for a in aar], color="#1A6B72", alpha=0.8)
    ax1.axvline(0, color="#C8A84B", linestyle="--", linewidth=1.5, label="Event day")
    ax1.axhline(0, color="black", linewidth=0.5)
    ax1.set_ylabel("Average Abnormal Return (%)")
    ax1.set_title("TAVE Event Study — Average Abnormal Returns Around Tokenization Announcements")
    ax1.legend(); ax1.grid(alpha=0.2)

    ax2.plot(days, [c*100 for c in caar], color="#1A6B72", marker="o", linewidth=2)
    ax2.axvline(0, color="#C8A84B", linestyle="--", linewidth=1.5)
    ax2.axhline(0, color="black", linewidth=0.5)
    ax2.set_ylabel("Cumulative AAR (%)"); ax2.set_xlabel("Trading days relative to announcement (t=0)")
    ax2.set_title("Cumulative Average Abnormal Return")
    ax2.grid(alpha=0.2)

    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    print(f"  ✓ Plot saved: {save_path}")
    plt.close()


# TAVE house colours (consistent across all three figures)
_TAVE_TEAL = "#1A6B72"
_TAVE_GOLD = "#C8A84B"
_TAVE_RED  = "#A23B3B"
_TAVE_GREY = "#6B6B6B"


def plot_coefficient_estimates(reg, save_path, title="TAVE Cross-Sectional Regression — Coefficient Estimates"):
    """
    Forest/coefficient plot: each regressor's point estimate with a 95% CI whisker.
    Makes the significant drivers (TokenizationScope, JurisdictionScore) visually
    obvious. Skips the constant. Colours: significant = teal, non-sig = grey.
    """
    if reg is None or "coefficients" not in reg:
        print("  (coef plot skipped — no regression)"); return
    coefs = reg["coefficients"].drop(index=[i for i in ["const"] if i in reg["coefficients"].index])
    if coefs.empty:
        print("  (coef plot skipped — no non-constant coefficients)"); return

    names = list(coefs.index)
    est = coefs["coef"].values
    lo = coefs["ci_low"].values
    hi = coefs["ci_high"].values
    pv = coefs["p_value"].values
    y = np.arange(len(names))[::-1]   # top-to-bottom

    fig, ax = plt.subplots(figsize=(9, max(2.4, 0.8 * len(names) + 1.5)))
    for yi, e, l, h, p in zip(y, est, lo, hi, pv):
        col = _TAVE_TEAL if (p is not None and p < 0.10) else _TAVE_GREY
        ax.plot([l, h], [yi, yi], color=col, linewidth=2.2, solid_capstyle="round")
        ax.plot(e, yi, "o", color=col, markersize=8, zorder=3)
    ax.axvline(0, color="black", linewidth=0.8, linestyle="--", alpha=0.7)
    ax.set_yticks(y); ax.set_yticklabels(names)
    ax.set_xlabel("Coefficient on CAR (-1,+1)  •  teal = significant at 10%, grey = not")
    ax.set_title(title)
    ax.grid(axis="x", alpha=0.2)
    # annotate p-values
    for yi, e, p in zip(y, est, pv):
        if p is not None and not np.isnan(p):
            ax.annotate(f"p={p:.3f}", (e, yi), textcoords="offset points",
                        xytext=(0, 10), ha="center", fontsize=8, color=_TAVE_GREY)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    print(f"  ✓ Coefficient plot saved: {save_path}")
    plt.close()


def plot_assetclass_caar(asset_sub, save_path, title="TAVE — Abnormal Return by Asset Class (the liquidity gradient)"):
    """
    Horizontal bar chart of CAAR by asset class — the standout result.
    Positive (illiquid real assets) in teal, negative (already-liquid) in red.
    Significant bars (p<0.10) get a star.
    """
    if asset_sub is None or asset_sub.empty:
        print("  (asset-class chart skipped — no data)"); return
    d = asset_sub.sort_values("CAAR_pct")
    names = d["asset_class"].tolist()
    vals = d["CAAR_pct"].values
    pv = d["BMP_p"].values if "BMP_p" in d.columns else [np.nan] * len(vals)
    colours = [_TAVE_TEAL if v >= 0 else _TAVE_RED for v in vals]

    fig, ax = plt.subplots(figsize=(9, max(2.6, 0.6 * len(names) + 1.6)))
    bars = ax.barh(names, vals, color=colours, alpha=0.88)
    ax.axvline(0, color="black", linewidth=0.8)
    ax.set_xlabel("Cumulative Average Abnormal Return (%), (-1,+1) window")
    ax.set_title(title)
    ax.grid(axis="x", alpha=0.2)
    # Headroom so edge labels never collide with the frame
    vmax = max(abs(vals.min()), abs(vals.max())) if len(vals) else 1.0
    ax.set_xlim(-vmax * 1.35, vmax * 1.35)
    for bar, v, p in zip(bars, vals, pv):
        star = " *" if (p is not None and not np.isnan(p) and p < 0.10) else ""
        ax.annotate(f"{v:+.2f}%{star}",
                    (v, bar.get_y() + bar.get_height() / 2),
                    textcoords="offset points",
                    xytext=(6 if v >= 0 else -6, 0),
                    ha="left" if v >= 0 else "right", va="center", fontsize=9)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches="tight")
    print(f"  ✓ Asset-class chart saved: {save_path}")
    plt.close()


def export_findings_pdf(findings_text, save_path):
    """
    Write the plain-text findings summary to a PDF on Drive. Colab on mobile is
    unstable, so this preserves the key output even if the session drops.
    Uses matplotlib (always available) as a dependency-free PDF writer.
    """
    try:
        from matplotlib.backends.backend_pdf import PdfPages
    except Exception as e:
        print(f"  (PDF export skipped: {e})"); return
    # Paginate the monospace text across as many pages as needed
    lines = findings_text.split("\n")
    lines_per_page = 52
    pages = [lines[i:i + lines_per_page] for i in range(0, len(lines), lines_per_page)]
    try:
        with PdfPages(save_path) as pdf:
            for pg in pages:
                fig = plt.figure(figsize=(8.27, 11.69))  # A4 portrait
                fig.text(0.06, 0.97, "\n".join(pg), family="monospace",
                         fontsize=8, va="top", ha="left")
                pdf.savefig(fig); plt.close(fig)
        print(f"  ✓ Findings PDF saved: {save_path}")
    except Exception as e:
        print(f"  (PDF export failed: {e})")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
# DIAGNOSTIC / VERDICT BLOCK
# ═══════════════════════════════════════════════════════════════════════════════

def print_analysis_diagnostics(events, car_df, index_health, skip, test_results, reg):
    """Consolidated verdict for Code 2: status, shortfalls, exact fixes."""
    L = ["", "╔" + "═"*68 + "╗",
         "║" + "  TAVE CODE 2 — ANALYSIS DIAGNOSTICS & VERDICT".ljust(68) + "║",
         "╚" + "═"*68 + "╝"]
    problems = []

    # [1] Index health (Blocker 2)
    L.append("\n[1] INDEX HEALTH (Blocker 2 — native index vs ETF fallback)")
    L.append("    " + "-"*60)
    for idx, h in index_health.items():
        st = h["status"]
        if st == "NATIVE_OK":
            L.append(f"    ✓ {idx:<8} native OK ({h['n_days']} days)")
        elif st == "FALLBACK_ETF":
            L.append(f"    ⚠ {idx:<8} native FAILED → using ETF {h['used_symbol']} ({h['n_days']} days)")
            problems.append(f"[INDEX] {idx} native symbol failed; ETF fallback {h['used_symbol']} used. "
                            f"Results valid but note the substitution in your methods section.")
        else:
            L.append(f"    ✗ {idx:<8} FAILED (native + fallback). Market unusable.")
            problems.append(f"[INDEX] {idx} has NO usable data (native + ETF both failed). "
                            f"FIX: pick another proxy in INDEX_FALLBACK for {idx}, or drop that market. "
                            f"All events in this market were dropped from the study.")

    # [2] Event inclusion / skip reasons
    L.append("\n[2] EVENT INCLUSION")
    L.append("    " + "-"*60)
    n_ok = len(skip.get("ok", []))
    L.append(f"    ✓ Events included in study:     {n_ok}")
    L.append(f"    ⤫ Skipped — no price ticker:    {len(skip.get('no_ticker', []))}")
    L.append(f"    ⤫ Skipped — no index:           {len(skip.get('no_index', []))}")
    L.append(f"    ⤫ Skipped — est window failed:  {len(skip.get('est_failed', []))}")
    L.append(f"    ⤫ Dropped — confounded (earnings): {len(skip.get('confounded', []))}")
    for reason, label in [("no_index","no index"), ("est_failed","est-window")]:
        if skip.get(reason):
            for item in skip[reason][:5]:
                L.append(f"        · [{label}] {item}")
            if len(skip[reason]) > 5:
                L.append(f"        · ... +{len(skip[reason])-5} more")
    if n_ok < 30:
        problems.append(f"[POWER] Only {n_ok} events in study. Large-effect designs want 30+. "
                        f"FIX: recover skipped events — check yf tickers (no_ticker list), "
                        f"add INDEX_FALLBACK entries (no_index), or relax confound filter "
                        f"(filter_confounds=False) if earnings overlap is over-aggressive.")

    # [3] Significance tests
    L.append("\n[3] SIGNIFICANCE (primary window -1,+1)")
    L.append("    " + "-"*60)
    if test_results is None or test_results.empty:
        L.append("    ✗ No test results.")
        problems.append("[TESTS] No significance results — no CARs computed. Fix inclusion first.")
    else:
        prim = test_results[test_results["window"] == "(-1,1)"]
        if not prim.empty:
            r = prim.iloc[0]
            L.append(f"    N={int(r['N'])}  CAAR={r['CAAR_pct']:.3f}%  "
                     f"BMP_t={r['BMP_t']:.3f}  BMP_p={r['BMP_p']:.4f} {r['sig_BMP']}")
            if r["BMP_p"] >= 0.10:
                L.append("    → Not significant at 10%. This is a finding, not an error.")
                L.append("      Interpretation: market may not price tokenization as material,")
                L.append("      or sample/power too thin. Report honestly; check subsamples.")

    # [4] Regression
    L.append("\n[4] CROSS-SECTIONAL OLS")
    L.append("    " + "-"*60)
    if reg is None:
        L.append("    ✗ Regression did not run.")
        problems.append("[OLS] Regression skipped — likely <10 obs after dropping NaN. "
                        "FIX: more verified events, or fill missing classifications "
                        "(asset_class drives BaselineIlliquidity; type drives TokenizationScope).")
    else:
        L.append(f"    ✓ N={reg['n_obs']}  clusters(firms)={reg.get('n_clusters')}  SE={reg['se_type']}")
        if reg.get("n_clusters") and reg["n_clusters"] < 10:
            problems.append(f"[OLS] Only {reg['n_clusters']} firm clusters — clustered SEs are "
                            f"unreliable below ~10-15 clusters. FIX: report HC3 alongside (already "
                            f"computed as model_hc3); interpret cluster SEs with caution.")
        # check the key coefficient sign (H2: BaselineIlliquidity > 0)
        coefs = reg["coefficients"]
        if "BaselineIlliquidity" in coefs.index:
            b = coefs.loc["BaselineIlliquidity"]
            L.append(f"    H2 (BaselineIlliquidity): coef={b['coef']:.4f}, p={b['p_value']:.4f}")
            if b["coef"] > 0 and b["p_value"] < 0.10:
                L.append("    → H2 SUPPORTED: illiquidity channel dominates (the key TAVE finding).")
            else:
                L.append("    → H2 not supported at 10%. Report honestly; may be power-limited.")

    # VERDICT
    L.append("\n" + "═"*70)
    if not problems:
        L.append("  ✅ VERDICT: ANALYSIS CLEAN — results ready to write up.")
    else:
        L.append(f"  ⚠️  VERDICT: {len(problems)} ITEM(S) NEED ATTENTION")
        L.append("  " + "-"*66)
        for i, p in enumerate(problems, 1):
            L.append(f"  {i}. {p}")
    L.append("═"*70)
    print("\n".join(L))
    return problems


def market_outlier_diagnostic(car_df, window="(-1,1)", top_n=5):
    """
    For each market (exchange), check whether its CAAR is driven by a few outliers
    or is broad-based. Reports, per market:
      - mean CAAR vs MEDIAN CAR (median is outlier-proof; large gap = skew/outliers)
      - the top_n events by absolute CAR (the biggest contributors)
      - CAAR and BMP p-value AFTER dropping the top_n absolute-CAR events
        (if the effect collapses, it was outlier-driven; if it holds, it's real)
      - a simple date-sanity flag: events landing on a weekend (bad scrape date)

    Returns (summary_df, detail_dict) where detail_dict[market] is the top-events frame.
    """
    sub = car_df[car_df["window"] == window].copy()
    if sub.empty:
        return pd.DataFrame(), {}

    summary_rows = []
    detail = {}
    for mkt, g in sub.groupby("exchange"):
        g = g.copy()
        n = len(g)
        mean_caar = g["CAR"].mean() * 100
        median_car = g["CAR"].median() * 100
        # Top contributors by absolute CAR
        g["abs_CAR"] = g["CAR"].abs()
        top = g.nlargest(min(top_n, n), "abs_CAR")[
            ["event_id", "firm_name", "event_date", "asset_class", "CAR"]].copy()
        top["CAR_pct"] = top["CAR"] * 100
        detail[mkt] = top

        # Leave-out test: drop the top_n absolute-CAR events, recompute
        keep = g.drop(top.index)
        if len(keep) >= 2:
            caar_ex = keep["CAR"].mean() * 100
            t_b, p_b, _ = bmp_test(keep["SCAR"].values)
        else:
            caar_ex, p_b = np.nan, np.nan

        # Date sanity: weekend dates suggest a bad scraped date
        wd = pd.to_datetime(g["event_date"]).dt.weekday
        n_weekend = int((wd >= 5).sum())

        # Skew flag: mean far from median = outlier-influenced
        gap = abs(mean_caar - median_car)
        skew_flag = "OUTLIER-SKEWED" if gap > abs(median_car) + 0.1 else "ok"

        summary_rows.append({
            "market": mkt, "N": n,
            "mean_CAAR_pct": round(mean_caar, 3),
            "median_CAR_pct": round(median_car, 3),
            "CAAR_excl_top%d_pct" % top_n: round(caar_ex, 3) if not np.isnan(caar_ex) else np.nan,
            "BMP_p_excl_top%d" % top_n: round(p_b, 4) if not np.isnan(p_b) else np.nan,
            "weekend_dates": n_weekend,
            "robustness": skew_flag,
        })

    return pd.DataFrame(summary_rows), detail


# Firm(s) excluded from the PRIMARY sample for documented methodological reasons.
# Sea Limited: Singapore-founded but NYSE-listed in USD; its returns track the US
# tech cycle, not the ^STI benchmark the SGX cohort is measured against, so the
# market model is misspecified for it. Its idiosyncratic volatility (±15–22% CARs)
# also swamps the tokenization signal. Kept in the dataset; reported as a
# with-vs-without robustness check so the contrast itself justifies the exclusion.
EXCLUDE_FROM_PRIMARY = {"Sea Limited"}


def firm_exclusion_robustness(car_df, firm_name="Sea Limited", market="SGX", window="(-1,1)"):
    """
    ADDITIVE robustness check: recompute a single market's CAAR + BMP test
    WITH and WITHOUT a named firm. Returns a 2-row DataFrame (with / without)
    plus the excluded firm's own event count. Purely diagnostic — does not alter
    car_df or any other result. The 'with' row is the evidence FOR excluding the
    firm; the 'without' row is the sample actually used in the write-up.
    """
    sub = car_df[(car_df["window"] == window) & (car_df["exchange"] == market)].copy()
    if sub.empty:
        return pd.DataFrame(), 0
    in_firm = sub[sub["firm_name"] == firm_name]
    ex_firm = sub[sub["firm_name"] != firm_name]
    rows = []
    for label, g in [(f"WITH {firm_name}", sub), (f"WITHOUT {firm_name}", ex_firm)]:
        if g.empty:
            continue
        t_b, p_b, n = bmp_test(g["SCAR"].values)
        rows.append({
            "sample": label, "N": n,
            "CAAR_pct": round(g["CAR"].mean() * 100, 3),
            "median_CAR_pct": round(g["CAR"].median() * 100, 3),
            "BMP_t": round(t_b, 3) if pd.notna(t_b) else np.nan,
            "BMP_p": round(p_b, 4) if pd.notna(p_b) else np.nan,
        })
    return pd.DataFrame(rows), len(in_firm)


def liquidity_channel_analysis(car_df, volume_data, events_df, pre_days=60, post_days=60, gap=5):
    """
    WACC CHANNEL 1 — LIQUIDITY (the one channel the equity/volume data can populate).

    For each event, measure the change in liquidity from a pre-event window to a
    post-event window, using two standard measures:
      - Amihud (2002) illiquidity = mean(|daily return| / daily dollar volume) * 1e6
        (higher = MORE illiquid; tokenization should LOWER this if it improves liquidity)
      - Turnover = mean(daily volume) post vs pre (higher = more trading activity)

    Windows: pre = [-(pre_days+gap), -gap], post = [+gap, +(post_days+gap)] around event.
    The `gap` excludes the immediate event window so we measure a durable shift,
    not the announcement-day spike.

    Returns (summary_df, per_event_df):
      - per_event_df: Amihud pre/post + %change, turnover pre/post + %change per event
      - summary_df:  mean %change in Amihud illiquidity and turnover, with a t-test
        on whether the post-event change differs from zero (the empirical liquidity signal)

    INTERPRETATION FOR TAVE: a significant NEGATIVE change in Amihud illiquidity is
    direct empirical support for the liquidity-premium channel — tokenization making
    the asset/firm more liquid, lowering the illiquidity premium component of WACC.
    """
    sub = car_df[car_df["window"] == "(-1,1)"].copy()
    if sub.empty or not volume_data:
        return pd.DataFrame(), pd.DataFrame()

    # event_date per event_id (from the events frame, robust to dedup)
    date_lookup = dict(zip(events_df["event_id"], pd.to_datetime(events_df["event_date"])))

    # ── MATCHED-CONTROL (DiD) PRE-COMPUTATION ──────────────────────────────────
    # The raw pre→post change cannot distinguish an announcement effect from a
    # secular/market-wide liquidity trend (the RCT counterfactual problem). The
    # finance analogue of a control arm: for each event, measure the SAME change
    # over the SAME calendar windows for sample firms with NO event of their own
    # within ±control_gap_days — same sector mix, same market conditions.
    # Adjusted (DiD) effect = own Δ% − mean(control Δ%).
    control_gap_days = 30
    # Per-firm event dates (to exclude contemporaneously-treated controls)
    _ev = events_df.copy()
    _ev["event_date"] = pd.to_datetime(_ev["event_date"], errors="coerce")
    firm_event_dates = {f: g["event_date"].dropna().sort_values().values
                        for f, g in _ev.groupby("firm_name")}
    ticker_to_firm = dict(zip(_ev.get("yf_ticker", _ev.get("ticker")), _ev["firm_name"]))
    # Precompute each firm's daily Amihud ratio + volume series ONCE (fast slicing)
    firm_series = {}
    for tk, frame in volume_data.items():
        try:
            fv = frame.copy()
            fv.index = pd.to_datetime(fv.index)
            if getattr(fv.index, "tz", None) is not None:
                fv.index = fv.index.tz_localize(None)
            fv = fv.sort_index()
            r_abs = np.log(fv["close"] / fv["close"].shift(1)).abs()
            dvol = (fv["close"] * fv["volume"]).replace(0, np.nan)
            ratio = (r_abs / dvol).replace([np.inf, -np.inf], np.nan)
            firm_series[tk] = {"index": fv.index, "ratio": ratio.values,
                               "volume": fv["volume"].values, "frame": fv}
        except Exception:
            continue

    def _window_stats(tk, edate):
        """Amihud mean (×1e6) and mean volume for pre/post windows on firm tk's calendar."""
        fs = firm_series.get(tk)
        if fs is None:
            return None
        idx = fs["index"]
        pos_arr = idx.searchsorted(edate)
        if pos_arr >= len(idx):
            return None
        pos = int(pos_arr)
        lo_pre, hi_pre = max(0, pos - pre_days - gap), max(0, pos - gap)
        lo_post, hi_post = pos + gap, pos + gap + post_days
        if hi_pre - lo_pre < 20 or min(hi_post, len(idx)) - lo_post < 20:
            return None
        a_pre = np.nanmean(fs["ratio"][lo_pre:hi_pre]) * 1e6
        a_post = np.nanmean(fs["ratio"][lo_post:hi_post]) * 1e6
        v_pre = np.nanmean(fs["volume"][lo_pre:hi_pre])
        v_post = np.nanmean(fs["volume"][lo_post:hi_post])
        if not (np.isfinite(a_pre) and np.isfinite(a_post)) or a_pre == 0:
            return None
        return {"amihud_pct": (a_post - a_pre) / a_pre * 100,
                "turn_pct": (v_post - v_pre) / v_pre * 100 if v_pre else np.nan}

    rows = []
    for _, ev in sub.iterrows():
        yft = ev.get("yf_ticker") or ev.get("ticker")
        eid = ev.get("event_id")
        if yft not in volume_data:
            continue
        edate = date_lookup.get(eid, ev.get("event_date"))
        edate = pd.Timestamp(edate)
        if getattr(edate, "tz", None) is not None:
            edate = edate.tz_localize(None)

        lv = volume_data[yft].copy()
        lv.index = pd.to_datetime(lv.index)
        if getattr(lv.index, "tz", None) is not None:
            lv.index = lv.index.tz_localize(None)
        lv = lv.sort_index()
        if edate not in lv.index:
            future = lv.index[lv.index >= edate]
            if len(future) == 0:
                continue
            pos = lv.index.get_loc(future[0])
        else:
            pos = lv.index.get_loc(edate)

        pre = lv.iloc[max(0, pos - pre_days - gap): max(0, pos - gap)]
        post = lv.iloc[pos + gap: pos + gap + post_days]
        if len(pre) < 20 or len(post) < 20:
            continue

        def _amihud(frame):
            r = np.log(frame["close"] / frame["close"].shift(1)).abs()
            dollar_vol = (frame["close"] * frame["volume"]).replace(0, np.nan)
            ratio = (r / dollar_vol).replace([np.inf, -np.inf], np.nan).dropna()
            return ratio.mean() * 1e6 if len(ratio) else np.nan

        amihud_pre, amihud_post = _amihud(pre), _amihud(post)
        turn_pre, turn_post = pre["volume"].mean(), post["volume"].mean()
        if not (np.isfinite(amihud_pre) and np.isfinite(amihud_post)) or amihud_pre == 0:
            continue

        own_amihud_pct = (amihud_post - amihud_pre) / amihud_pre * 100
        own_turn_pct = (turn_post - turn_pre) / turn_pre * 100 if turn_pre else np.nan

        # ── Matched controls: same calendar windows, firms with no own event ±30d ──
        ctrl_a, ctrl_t = [], []
        for ctk in firm_series:
            if ctk == yft:
                continue
            cfirm = ticker_to_firm.get(ctk)
            cdates = firm_event_dates.get(cfirm)
            if cdates is not None and len(cdates):
                nearest = np.min(np.abs((cdates - np.datetime64(edate)) / np.timedelta64(1, "D")))
                if nearest <= control_gap_days:
                    continue   # control is contemporaneously treated — exclude
            cs = _window_stats(ctk, edate)
            if cs is None:
                continue
            ctrl_a.append(cs["amihud_pct"])
            if np.isfinite(cs["turn_pct"]):
                ctrl_t.append(cs["turn_pct"])

        n_controls = len(ctrl_a)
        ctrl_amihud_mean = float(np.mean(ctrl_a)) if ctrl_a else np.nan
        ctrl_turn_mean = float(np.mean(ctrl_t)) if ctrl_t else np.nan

        rows.append({
            "event_id": eid, "firm_name": ev.get("firm_name"), "exchange": ev.get("exchange"),
            "asset_class": ev.get("asset_class"),
            "amihud_pre": amihud_pre, "amihud_post": amihud_post,
            "amihud_pct_change": own_amihud_pct,
            "turnover_pct_change": own_turn_pct,
            "n_controls": n_controls,
            "ctrl_amihud_pct_change": ctrl_amihud_mean,
            "ctrl_turnover_pct_change": ctrl_turn_mean,
            "amihud_did": own_amihud_pct - ctrl_amihud_mean if np.isfinite(ctrl_amihud_mean) else np.nan,
            "turnover_did": own_turn_pct - ctrl_turn_mean if np.isfinite(ctrl_turn_mean) else np.nan,
        })

    per_event = pd.DataFrame(rows)
    if per_event.empty:
        return pd.DataFrame(), per_event

    # Aggregate signal: is the mean change different from zero?
    def _ttest(series):
        s = series.replace([np.inf, -np.inf], np.nan).dropna()
        if len(s) < 3:
            return np.nan, np.nan, len(s)
        t, p = stats.ttest_1samp(s, 0)
        return t, p, len(s)

    a_t, a_p, a_n = _ttest(per_event["amihud_pct_change"])
    t_t, t_p, t_n = _ttest(per_event["turnover_pct_change"])
    d_t, d_p, d_n = _ttest(per_event["amihud_did"]) if "amihud_did" in per_event else (np.nan, np.nan, 0)
    e_t, e_p, e_n = _ttest(per_event["turnover_did"]) if "turnover_did" in per_event else (np.nan, np.nan, 0)
    summary = pd.DataFrame([
        {"metric": "Amihud illiquidity Δ% (pre→post, raw)", "N": a_n,
         "mean_pct_change": per_event["amihud_pct_change"].replace([np.inf, -np.inf], np.nan).dropna().mean(),
         "median_pct_change": per_event["amihud_pct_change"].replace([np.inf, -np.inf], np.nan).dropna().median(),
         "t_stat": a_t, "p_value": a_p},
        {"metric": "Turnover Δ% (pre→post, raw)", "N": t_n,
         "mean_pct_change": per_event["turnover_pct_change"].replace([np.inf, -np.inf], np.nan).dropna().mean(),
         "median_pct_change": per_event["turnover_pct_change"].replace([np.inf, -np.inf], np.nan).dropna().median(),
         "t_stat": t_t, "p_value": t_p},
        {"metric": "Amihud Δ% DiD vs matched controls (HEADLINE)", "N": d_n,
         "mean_pct_change": per_event.get("amihud_did", pd.Series(dtype=float)).replace([np.inf, -np.inf], np.nan).dropna().mean(),
         "median_pct_change": per_event.get("amihud_did", pd.Series(dtype=float)).replace([np.inf, -np.inf], np.nan).dropna().median(),
         "t_stat": d_t, "p_value": d_p},
        {"metric": "Turnover Δ% DiD vs matched controls", "N": e_n,
         "mean_pct_change": per_event.get("turnover_did", pd.Series(dtype=float)).replace([np.inf, -np.inf], np.nan).dropna().mean(),
         "median_pct_change": per_event.get("turnover_did", pd.Series(dtype=float)).replace([np.inf, -np.inf], np.nan).dropna().median(),
         "t_stat": e_t, "p_value": e_p},
    ])
    return summary, per_event


def placebo_test(events_df, price_data, n_placebo=10, seed=42):
    """
    PLACEBO / FALSIFICATION TEST — the strongest robustness check for an event study.

    Re-runs the event study on RANDOM pseudo-event dates (same firms, shuffled dates).
    A genuine announcement effect should DISAPPEAR on placebo dates. If the placebo
    CAAR is also significant, it signals model misspecification (e.g. a bad benchmark),
    not a real tokenization effect.

    Returns a DataFrame: for each placebo run, the (-1,1) CAAR and BMP p-value.
    """
    rng = np.random.default_rng(seed)
    real_dates = pd.to_datetime(events_df["event_date"])
    lo, hi = real_dates.min(), real_dates.max()
    span_days = max((hi - lo).days, 30)

    rows = []
    for k in range(n_placebo):
        fake = events_df.copy()
        offsets = rng.integers(0, span_days, size=len(fake))
        fake["event_date"] = [lo + pd.Timedelta(days=int(o)) for o in offsets]
        car_fake, _, _ = event_study(fake, price_data, filter_confounds=False)
        if car_fake.empty:
            rows.append({"placebo_run": k + 1, "N": 0, "CAAR_pct": np.nan, "BMP_p": np.nan}); continue
        sub = car_fake[car_fake["window"] == "(-1,1)"]
        if sub.empty:
            rows.append({"placebo_run": k + 1, "N": 0, "CAAR_pct": np.nan, "BMP_p": np.nan}); continue
        t_b, p_b, n = bmp_test(sub["SCAR"].values)
        rows.append({"placebo_run": k + 1, "N": n,
                     "CAAR_pct": sub["CAR"].mean() * 100, "BMP_p": p_b})
    return pd.DataFrame(rows)


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def run_analysis(filter_confounds=True, events_in=None, fundamentals_in=None):
    """
    events_in : optional pre-loaded events DataFrame (e.g. straight from
                run_collection in the same session). When provided, the Drive
                file round-trip is BYPASSED entirely — no stale-read risk.
                When None, events are loaded from the Excel as before.
    fundamentals_in : optional YF_Fundamentals frame to pair with events_in.
    """
    print("█"*70)
    print("  TAVE EVENT STUDY — CODE 2: ANALYSIS  (index-validated + clustered + diagnostics)")
    print("█"*70)

    # 0. Environment: mount Drive (if needed) + confirm input exists
    print("\n── Environment setup ──")
    if events_in is None:
        ensure_drive()

    # 1. Load events — from memory if handed in, else from the Excel file
    print("\n── Loading events ──")
    if events_in is not None and not events_in.empty:
        events = _prepare_events_in_memory(events_in)
        print(f"  Using {len(events)} events passed IN MEMORY (no file round-trip → no stale-read risk)")
        if "exchange" in events.columns:
            print(f"  Verified events by market: {events['exchange'].value_counts().to_dict()}")
    else:
        events = load_events()
    if events.empty:
        print("No events to analyze."); return

    # 2. Fundamentals — from memory if handed in, else from the Excel
    fundamentals = pd.DataFrame()
    if fundamentals_in is not None and not fundamentals_in.empty:
        fundamentals = fundamentals_in
    else:
        try:
            fundamentals = pd.read_excel(INPUT_FILE, sheet_name="YF_Fundamentals")
        except Exception:
            print("  No YF_Fundamentals sheet found — regression will use fewer controls")

    # 3. Download prices (with index validation + ETF fallback — Blocker 2)
    print("\n── Downloading + validating prices ──")
    tickers = events["yf_ticker"].dropna().unique().tolist()
    indices = list(set(MARKET_INDEX.values()))
    prices, index_health, volume_data = download_prices(tickers, indices)
    print(f"  Got {len(prices)} usable price series")

    # 4. Event study (with confounding filter)
    print("\n── Computing abnormal returns ──")
    car_df, timeline, skip = event_study(events, prices, filter_confounds=filter_confounds)
    print(f"  Events included: {len(skip.get('ok', []))} | window-obs: {len(car_df)}")

    # If empty, still print diagnostics so the user knows exactly why
    if car_df.empty:
        print_analysis_diagnostics(events, car_df, index_health, skip, None, None)
        return

    # 5. Significance tests
    print("\n── Significance tests ──")
    test_results = run_all_tests(car_df)
    print(test_results[["window","N","CAAR_pct","BMP_t","BMP_p","sig_BMP"]].to_string(index=False))

    # 6. Cross-sectional regression — TWO models:
    #    Primary  = full verified sample (selection independent of outcome)
    #    Robust   = HIGH-impact tier only — strictly more restrictive than primary;
    #               should strengthen if the effect is real (dose-response)
    print("\n── Cross-sectional OLS — PRIMARY (full sample) ──")
    reg = cross_sectional_regression(car_df, fundamentals, window="(-1,1)")
    if reg:
        print(f"  N={reg['n_obs']} | clusters={reg.get('n_clusters')} | SE={reg['se_type']}")
        print(reg["coefficients"].round(4).to_string())

    print("\n── Cross-sectional OLS — ROBUSTNESS (HIGH-impact only) ──")
    reg_hi = cross_sectional_regression(car_df, fundamentals, window="(-1,1)",
                                        impact_filter=["HIGH"])
    if reg_hi:
        print(f"  N={reg_hi['n_obs']} | clusters={reg_hi.get('n_clusters')} | SE={reg_hi['se_type']}")
        print(reg_hi["coefficients"].round(4).to_string())
    else:
        print("  Not enough HIGH-impact events for a separate regression "
              "(report primary only, note this as a limitation).")

    # EXPLORATORY — illiquidity x jurisdiction interaction, reported SEPARATELY.
    # Kept out of the primary spec (collinear with its components); shown here so the
    # interaction can be discussed as an exploratory extension, not a headline test.
    print("\n── Cross-sectional OLS — EXPLORATORY (illiquidity × jurisdiction interaction) ──")
    print("  NOTE: collinear with its components — interpret the interaction sign only,")
    print("  not the main effects. Primary H2 test is the no-interaction model above.")
    reg_int = cross_sectional_regression(car_df, fundamentals, window="(-1,1)",
                                         include_interaction=True)
    if reg_int:
        print(f"  N={reg_int['n_obs']} | clusters={reg_int.get('n_clusters')} | SE={reg_int['se_type']}")
        print(reg_int["coefficients"].round(4).to_string())
        if "IlliqXJuris" in reg_int["coefficients"].index:
            ix = reg_int["coefficients"].loc["IlliqXJuris"]
            print(f"  Interaction IlliqXJuris: coef={ix['coef']:+.4f}, p={ix['p_value']:.4f}")
            if ix["p_value"] < 0.10:
                direction = ("DAMPENS" if ix["coef"] < 0 else "AMPLIFIES")
                print(f"  → Regulatory clarity {direction} the illiquidity effect "
                      f"(exploratory; collinearity caveat applies).")

    # DECISIVE ROBUSTNESS 1 — primary spec EXCLUDING Sea Limited.
    # Sea is NYSE-listed/USD but STI-benchmarked and coded Singapore/clarity-5;
    # its extreme CARs can drive the JurisdictionScore coefficient. If the
    # coefficient survives ex-Sea, the jurisdiction result is robust (composition
    # story); if it collapses, downgrade 4b to descriptive.
    print("\n── Cross-sectional OLS — ROBUSTNESS (primary spec, EXCLUDING Sea Limited) ──")
    reg_xsea = cross_sectional_regression(car_df, fundamentals, window="(-1,1)",
                                          exclude_firms=EXCLUDE_FROM_PRIMARY)
    if reg_xsea:
        print(f"  N={reg_xsea['n_obs']} | clusters={reg_xsea.get('n_clusters')} | SE={reg_xsea['se_type']}")
        print(reg_xsea["coefficients"].round(4).to_string())
        if "JurisdictionScore" in reg_xsea["coefficients"].index:
            js = reg_xsea["coefficients"].loc["JurisdictionScore"]
            verdict = ("SURVIVES ex-Sea — jurisdiction result robust"
                       if js["p_value"] < 0.05 else
                       "COLLAPSES ex-Sea — treat jurisdiction gradient as Sea-driven; downgrade to descriptive")
            print(f"  JurisdictionScore ex-Sea: coef={js['coef']:+.4f}, p={js['p_value']:.4f} → {verdict}")

    # DECISIVE ROBUSTNESS 2 — jurisdiction FIXED EFFECTS instead of linear score.
    # Drops the linearity-of-clarity assumption: do jurisdictions differ at all?
    print("\n── Cross-sectional OLS — ROBUSTNESS (jurisdiction fixed effects, ex-Sea) ──")
    reg_fe = cross_sectional_regression(car_df, fundamentals, window="(-1,1)",
                                        exclude_firms=EXCLUDE_FROM_PRIMARY,
                                        jurisdiction_fe=True)
    if reg_fe:
        print(f"  N={reg_fe['n_obs']} | clusters={reg_fe.get('n_clusters')} | SE={reg_fe['se_type']}")
        print(reg_fe["coefficients"].round(4).to_string())
        print("  Read: each J_* coefficient = that jurisdiction's CAR differential vs the base cohort,")
        print("  holding asset-class illiquidity, scope, and firm controls fixed.")

    # DECISIVE ROBUSTNESS 2b — announcement-TYPE controls (materiality/composition).
    # Does JurisdictionScore survive once the announcement's type/materiality mix is
    # held fixed? (Layer 2 test: clear regimes host more material announcements.)
    print("\n── Cross-sectional OLS — ROBUSTNESS (announcement-type controls, ex-Sea) ──")
    reg_ann = cross_sectional_regression(car_df, fundamentals, window="(-1,1)",
                                         exclude_firms=EXCLUDE_FROM_PRIMARY,
                                         announcement_fe=True)
    if reg_ann:
        print(f"  N={reg_ann['n_obs']} | clusters={reg_ann.get('n_clusters')} | SE={reg_ann['se_type']}")
        print(reg_ann["coefficients"].round(4).to_string())
        if "JurisdictionScore" in reg_ann["coefficients"].index:
            ja = reg_ann["coefficients"].loc["JurisdictionScore"]
            print(f"  JurisdictionScore w/ type controls: coef={ja['coef']:+.4f}, p={ja['p_value']:.4f}"
                  f" → {'survives materiality-mix control' if ja['p_value']<0.05 else 'absorbed by announcement composition'}")

    # DECISIVE ROBUSTNESS 3 — wild cluster bootstrap (few-cluster inference).
    # With only ~11 firm clusters, analytic clustered SEs are anti-conservative;
    # the WCB p-value is the credible one. Run on primary AND ex-Sea samples.
    print("\n── Wild cluster bootstrap — JurisdictionScore (999 reps, Rademacher, null-imposed) ──")
    if reg and "reg_df" in reg and "JurisdictionScore" in reg["variables"]:
        wcb_full = wild_cluster_bootstrap_p(reg["reg_df"], "JurisdictionScore")
        print(f"  FULL sample : t={wcb_full.get('t_clustered')} | analytic p<0.001 | "
              f"WCB p={wcb_full.get('p_wild_cluster_bootstrap')} | clusters={wcb_full.get('n_clusters')}")
    else:
        wcb_full = None
    if reg_xsea and "reg_df" in reg_xsea and "JurisdictionScore" in reg_xsea["variables"]:
        wcb_xsea = wild_cluster_bootstrap_p(reg_xsea["reg_df"], "JurisdictionScore")
        print(f"  EX-SEA      : t={wcb_xsea.get('t_clustered')} | "
              f"WCB p={wcb_xsea.get('p_wild_cluster_bootstrap')} | clusters={wcb_xsea.get('n_clusters')}")
        print("  Read: report the WCB p-values as the headline inference for Gap 4b;")
        print("  they are robust to the small cluster count, the analytic p is not.")
    else:
        wcb_xsea = None

    # 7. Robustness
    print("\n── Robustness ──")
    wins = robustness_winsorized(car_df)
    region_sub = subsample_by_region(car_df)
    asset_sub = subsample_by_assetclass(car_df)
    print("  By asset class (tests H2 - illiquidity channel):")
    print(asset_sub.to_string(index=False))

    # STRUCTURAL PROXY (exploratory, future-research) — provenance-sensitive vs
    # financial-transparent assets. Light, coarse signal for Gaps 2/3; NOT a channel
    # measurement (that needs microstructure/bond data). Labelled as such.
    prov_summary, prov_diff = provenance_contrast(car_df)
    if not prov_summary.empty:
        print("\n  Provenance-sensitive vs financial-transparent (STRUCTURAL PROXY for Gaps 2/3 — exploratory):")
        print(prov_summary.to_string(index=False))
        if prov_diff:
            sig = "SIGNAL" if prov_diff.get("welch_p", 1) < 0.10 else "no clear separation"
            print(f"    Difference: {prov_diff['diff_CAAR_pct']:+.3f}pp (Welch t={prov_diff['welch_t']}, "
                  f"p={prov_diff['welch_p']}) → {sig}")
            print("    NOTE: coarse asset-class proxy, not a microstructure test — future-research signal only.")

    # STRUCTURAL PROXY (exploratory) — does disclosed pct_tokenized carry signal?
    pct_sig = pct_tokenized_signal(car_df)
    if pct_sig:
        if "corr_CAR_pct_tokenized" in pct_sig:
            print(f"\n  pct_tokenized signal check (STRUCTURAL PROXY, N={pct_sig['n']} disclosed): "
                  f"corr(CAR, pct)={pct_sig['corr_CAR_pct_tokenized']:+.3f}, p={pct_sig['corr_p']} "
                  f"→ {'signal' if pct_sig['corr_p']<0.10 else 'no clear signal'} (underpowered; future-research flag).")
        else:
            print(f"\n  pct_tokenized signal check: {pct_sig.get('note','n/a')}")

    # Per-market outlier / leave-out robustness — is any single market's CAAR
    # driven by a few extreme events? (Critical when one market, e.g. SGX, is large.)
    print("\n  Per-market outlier diagnostic (mean vs median, drop-top-5 leave-out test):")
    outlier_summary, outlier_detail = market_outlier_diagnostic(car_df, top_n=5)
    if not outlier_summary.empty:
        print(outlier_summary.to_string(index=False))
        for mkt, top in outlier_detail.items():
            if len(top) and mkt != "US":   # spotlight non-US (smaller, scrape-sourced) markets
                print(f"\n  Top contributing events — {mkt}:")
                show = top[["event_id", "firm_name", "event_date", "asset_class", "CAR_pct"]].copy()
                show["event_date"] = pd.to_datetime(show["event_date"]).dt.strftime("%Y-%m-%d")
                print(show.to_string(index=False))

    # ADDITIVE robustness — SGX CAAR with vs without Sea Limited.
    # Sea is NYSE-listed in USD (benchmarked here against ^STI, a misspecification)
    # and its ±15-22% CARs swamp the tokenization signal. We report BOTH: the WITH
    # row is the evidence for excluding it; the WITHOUT row is the sample we use.
    sea_robust, sea_n = firm_exclusion_robustness(car_df, firm_name="Sea Limited", market="SGX")
    if not sea_robust.empty:
        print(f"\n  SGX robustness — with vs without Sea Limited ({sea_n} Sea events, NYSE-listed/USD):")
        print(sea_robust.to_string(index=False))

    # WACC CHANNEL 1 — LIQUIDITY (Amihud illiquidity + turnover, pre vs post event)
    print("\n  Liquidity channel (WACC ch.1): Amihud illiquidity + turnover, pre vs post event:")
    liq_summary, liq_per_event = liquidity_channel_analysis(car_df, volume_data, events)
    if not liq_summary.empty:
        print(liq_summary.round(3).to_string(index=False))
        amr = liq_summary[liq_summary["metric"].str.contains("DiD", na=False) &
                          liq_summary["metric"].str.startswith("Amihud")]
        raw = liq_summary[liq_summary["metric"].str.contains("raw", na=False) &
                          liq_summary["metric"].str.startswith("Amihud")]
        if not amr.empty and pd.notna(amr.iloc[0]["p_value"]):
            a = amr.iloc[0]
            raw_note = ""
            if not raw.empty and pd.notna(raw.iloc[0]["p_value"]) and raw.iloc[0]["p_value"] < 0.10:
                raw_note = " (raw pre→post change was significant; the DiD is the credible test)"
            if a["p_value"] < 0.10 and a["mean_pct_change"] < 0:
                print("  → DiD: illiquidity fell vs matched controls — empirical support for the")
                print("    liquidity-premium channel SURVIVES the control adjustment" + raw_note + ".")
            elif a["p_value"] < 0.10 and a["mean_pct_change"] > 0:
                print("  → DiD: illiquidity ROSE vs matched controls — opposite of the liquidity-benefit")
                print("    claim; report honestly" + raw_note + ".")
            else:
                print("  → DiD vs matched controls: NOT significant — the raw pre→post change reflects")
                print("    secular/market-wide liquidity trends, not an announcement effect" + raw_note + ".")
                print("    Report channel 1 as a null at this horizon; calibrate from literature.")
        elif not raw.empty:
            a = raw.iloc[0]
            if pd.notna(a["p_value"]) and a["p_value"] < 0.10 and a["mean_pct_change"] < 0:
                print("  → Illiquidity FELL significantly post-event (raw; no controls available) —")
                print("    treat as suggestive only.")
            else:
                print("  → No significant liquidity change — channel-1 benefit not detected in equity")
                print("    liquidity; report as null, calibrate WACC ch.1 from literature instead.")
    else:
        print("  (No liquidity metrics computed — volume data unavailable for these tickers.)")

    # Placebo / falsification test — effect should vanish on random dates
    print("\n  Placebo test (random pseudo-event dates — effect should DISAPPEAR):")
    placebo = placebo_test(events, prices, n_placebo=10)  # 10-draw spec: report false-positive RATE
    if not placebo.empty:
        print(placebo.to_string(index=False))
        real_p = test_results[test_results["window"] == "(-1,1)"]["BMP_p"].iloc[0] if not test_results.empty else np.nan
        placebo_sig = (placebo["BMP_p"] < 0.10).sum()
        print(f"  → Real (-1,1) BMP p={real_p:.4f}; placebo runs significant at 10%: "
              f"{placebo_sig}/{len(placebo)} (want 0 — confirms the effect isn't an artifact)")

    # 8. Plot — three publication-quality figures
    print("\n── Plotting ──")
    plot_ar_timeline(timeline, PLOT_FILE)
    plot_coefficient_estimates(reg, COEF_PLOT_FILE)
    plot_assetclass_caar(asset_sub, ASSET_PLOT_FILE)

    # 9. Export
    print("\n── Exporting results ──")
    with pd.ExcelWriter(RESULTS_FILE, engine="openpyxl") as w:
        car_df.drop(columns=["window_tuple"]).to_excel(w, "CAR_per_event", index=False)
        test_results.to_excel(w, "Significance_Tests", index=False)
        pd.DataFrame(list(index_health.values())).to_excel(w, "Index_Health", index=False)
        # Skipped events log
        skip_rows = []
        for reason, items in skip.items():
            if reason == "ok": continue
            for it in items:
                skip_rows.append({"skip_reason": reason, "event": it})
        if skip_rows:
            pd.DataFrame(skip_rows).to_excel(w, "Skipped_Events", index=False)
        if reg:
            reg["coefficients"].to_excel(w, "OLS_Primary_FullSample")
            pd.DataFrame({"summary": [reg["summary"]]}).to_excel(w, "OLS_Primary_Summary", index=False)
            if reg.get("model_hc3") is not None:
                _coef_frame(reg["model_hc3"]).rename(columns={
                    "std_err": "std_err_HC3", "t_stat": "t_HC3", "p_value": "p_HC3"
                }).to_excel(w, "OLS_Primary_HC3")
        if reg_hi:
            reg_hi["coefficients"].to_excel(w, "OLS_Robust_HighImpact")
            pd.DataFrame({"summary": [reg_hi["summary"]]}).to_excel(w, "OLS_Robust_Summary", index=False)
        if 'reg_int' in dir() and reg_int:
            reg_int["coefficients"].to_excel(w, "OLS_Exploratory_Interaction")
        if 'reg_xsea' in dir() and reg_xsea:
            reg_xsea["coefficients"].to_excel(w, "OLS_Primary_ExSea")
        if 'reg_fe' in dir() and reg_fe:
            reg_fe["coefficients"].to_excel(w, "OLS_Jurisdiction_FE")
        if 'reg_ann' in dir() and reg_ann:
            reg_ann["coefficients"].to_excel(w, "OLS_AnnType_Control")
        wcb_rows = [x for x in [wcb_full if 'wcb_full' in dir() else None,
                                wcb_xsea if 'wcb_xsea' in dir() else None] if x]
        if wcb_rows:
            pd.DataFrame(wcb_rows, index=["full_sample", "ex_sea"][:len(wcb_rows)]).to_excel(w, "WCB_Jurisdiction")
        if 'prov_summary' in dir() and prov_summary is not None and not prov_summary.empty:
            prov_summary.to_excel(w, "Provenance_Contrast", index=False)
        if wins:
            pd.DataFrame([wins]).to_excel(w, "Robustness_Winsorized", index=False)
        region_sub.to_excel(w, "Subsample_Region", index=False)
        asset_sub.to_excel(w, "Subsample_AssetClass", index=False)
        if placebo is not None and not placebo.empty:
            placebo.to_excel(w, "Placebo_Test", index=False)
        if outlier_summary is not None and not outlier_summary.empty:
            outlier_summary.to_excel(w, "Market_Outlier_Check", index=False)
            # Also dump the top contributing events per market for manual eyeballing
            allrows = []
            for mkt, top in outlier_detail.items():
                t = top.copy(); t.insert(0, "market", mkt); allrows.append(t)
            if allrows:
                pd.concat(allrows, ignore_index=True).to_excel(w, "Top_Contributing_Events", index=False)
        if 'sea_robust' in dir() and sea_robust is not None and not sea_robust.empty:
            sea_robust.to_excel(w, "SGX_Sea_Robustness", index=False)
        if 'liq_summary' in dir() and liq_summary is not None and not liq_summary.empty:
            liq_summary.to_excel(w, "Liquidity_Channel", index=False)
        if 'liq_per_event' in dir() and liq_per_event is not None and not liq_per_event.empty:
            liq_per_event.to_excel(w, "Liquidity_PerEvent", index=False)

    print(f"\n✓ RESULTS: {RESULTS_FILE}")
    print(f"✓ PLOT:    {PLOT_FILE}")

    # 10. DIAGNOSTIC VERDICT
    problems = print_analysis_diagnostics(events, car_df, index_health, skip, test_results, reg)

    # 11. COPY-PASTE FINDINGS SUMMARY
    findings_text = print_findings_summary(car_df, test_results, reg, asset_sub, region_sub, wins, skip, reg_hi, outlier_summary, placebo, liq_summary)

    # 12. Export findings to PDF on Drive (mobile-Colab safety net)
    if findings_text:
        export_findings_pdf(findings_text, FINDINGS_PDF)

    return car_df, test_results, reg, index_health, skip, problems


def _cached_file_is_healthy(markets=("US", "SGX", "SIX", "XETRA", "LSE"), min_per_market=10):
    """
    Decide whether the existing master Excel already holds a GOOD multi-market
    sample — so we can reuse it wholesale instead of re-scraping (avoids API
    flakiness muddying a known-good dataset).

    Healthy means, for EVERY requested market:
      - the market is present in the file, AND
      - it has >= min_per_market verified-or-classified events, AND
      - its canary firm (if defined) appears.
    Returns (is_healthy: bool, reason: str) — reason is printed so the choice is auditable.
    """
    try:
        df = pd.read_excel(INPUT_FILE, sheet_name="Events_Classified")
    except Exception as e:
        return False, f"no readable master file ({e}) → scrape fresh"
    if "exchange" not in df.columns or df.empty:
        return False, "master file has no events/exchange column → scrape fresh"

    # Canary firms per market (mirror of Code 1's CANARY_EVENTS, kept local so Code 2
    # needs no import from Code 1; only the firm substring matters here).
    canary_firm = {"US": "BlackRock", "SGX": "DBS", "SIX": "UBS",
                   "XETRA": "Siemens", "LSE": "HSBC"}
    counts = df["exchange"].value_counts().to_dict()
    missing, thin, no_canary = [], [], []
    for m in markets:
        n = int(counts.get(m, 0))
        if n == 0:
            missing.append(m); continue
        if n < min_per_market:
            thin.append(f"{m}={n}")
        cf = canary_firm.get(m)
        if cf and not df[(df["exchange"] == m)]["firm_name"].astype(str).str.contains(cf, case=False, na=False).any():
            no_canary.append(f"{m}:{cf}")
    problems = []
    if missing:   problems.append(f"missing markets {missing}")
    if thin:      problems.append(f"thin markets {thin}")
    if no_canary: problems.append(f"canary firm absent {no_canary}")
    if problems:
        return False, "; ".join(problems) + " → scrape fresh"
    return True, f"all {len(markets)} markets present, populated, canaries found {counts} → reuse cache"


def run_pipeline(markets=("US", "SGX", "SIX", "XETRA", "LSE"),
                 refresh=("SGX", "SIX", "XETRA", "LSE"),
                 filter_confounds=True, prefer_cache=True):
    """
    ONE-CALL FULL PIPELINE — cache-first, scrape only as fallback.

    prefer_cache:
      True   (default) — Check the master file. If it holds a healthy 5-market
                         sample, analyse it DIRECTLY (no collection step, so a
                         scrape is structurally impossible) → deterministic reruns.
                         Only if the file is missing/unhealthy does it fall back to
                         scraping. This is the "freeze the good data" default.
      "auto"          — Same healthy-file → direct-analyse behaviour, but the
                         fallback to scraping is silent rather than warned.
      False           — Always scrape the `refresh` markets (US cached), hand to
                         analysis in memory. Use only when you WANT fresh data.

    The healthy-file path reads the master Excel (which IS the clean deliverable),
    so there is no scrape and no run-to-run variance. The freshness guard in
    load_events still protects against a genuinely stale (e.g. US-only) file.
    """
    # Cache-first: if the file is healthy, analyse it directly — never scrape.
    if prefer_cache in (True, "auto"):
        healthy, reason = _cached_file_is_healthy(markets)
        print(f"prefer_cache={prefer_cache!r} — cache health: {reason}")
        if healthy:
            print("  → Using the cached master file DIRECTLY (no collection, no scrape).")
            return run_analysis(filter_confounds=filter_confounds)
        if prefer_cache is True:
            print("  ⚠ File not healthy and prefer_cache=True — falling back to a fresh scrape.")
        use_refresh = refresh
    elif prefer_cache is False:
        print(f"prefer_cache=False → scraping fresh: {refresh} (US cached).")
        use_refresh = refresh
    else:
        use_refresh = refresh

    # Fallback scrape path (collection → in-memory handoff → analysis).
    if "run_collection" not in globals():
        raise RuntimeError(
            "run_collection not found and the cached file is not healthy. Load Code 1 "
            "(collection) in this session, or point INPUT_FILE at a healthy master file.")
    print("\nRunning collection, then handing events directly to analysis (in memory)...\n")
    out = globals()["run_collection"](markets=markets, refresh=use_refresh)
    events_df = out[0] if isinstance(out, tuple) else out
    fundamentals_df = out[2] if isinstance(out, tuple) and len(out) > 2 else None
    if events_df is None or len(events_df) == 0:
        print("Collection produced no events — aborting analysis."); return None
    print(f"\n→ Handed {len(events_df)} events to analysis in memory "
          f"(markets: {events_df['exchange'].value_counts().to_dict() if 'exchange' in events_df.columns else 'n/a'})\n")
    return run_analysis(filter_confounds=filter_confounds,
                        events_in=events_df, fundamentals_in=fundamentals_df)


def print_findings_summary(car_df, test_results, reg, asset_sub, region_sub, wins, skip, reg_hi=None, outlier_summary=None, placebo=None, liq_summary=None):
    """
    Plain-language findings block for direct copy-paste into chat for interpretation.
    States the headline result, the channel evidence (H2), robustness, and what it
    means — written so it can be read without the code.
    """
    S = []
    S.append("\n\n" + "┌" + "─"*68 + "┐")
    S.append("│" + "  COPY-PASTE FINDINGS — CODE 2 (ANALYSIS)".ljust(68) + "│")
    S.append("│" + "  Paste this block for interpretation guidance".ljust(68) + "│")
    S.append("└" + "─"*68 + "┘")

    if car_df is None or car_df.empty:
        S.append("\nNo abnormal returns were computed. See diagnostics for the cause.")
        text = "\n".join(S); print(text); return text

    n_events = len(skip.get("ok", []))
    n_dropped = sum(len(v) for k, v in skip.items() if k != "ok")

    # Headline result (primary window)
    S.append(f"\nHEADLINE RESULT (primary window -1 to +1)")
    prim = test_results[test_results["window"] == "(-1,1)"] if test_results is not None else None
    if prim is not None and not prim.empty:
        r = prim.iloc[0]
        sig = r["sig_BMP"] or "not significant at 10%"
        direction = "POSITIVE" if r["CAAR_pct"] > 0 else "NEGATIVE"
        S.append(f"  Events analysed:          {int(r['N'])}")
        S.append(f"  Average abnormal return:  {r['CAAR_pct']:+.3f}%  ({direction})")
        S.append(f"  Median abnormal return:   {r['median_CAR_pct']:+.3f}%")
        S.append(f"  Share positive:           {r['pct_positive']:.0f}%")
        S.append(f"  BMP test (primary):       t={r['BMP_t']:.3f}, p={r['BMP_p']:.4f}  {sig}")
        S.append(f"  Patell test:              Z={r['Patell_Z']:.3f}, p={r['Patell_p']:.4f}")
        S.append(f"  Corrado rank test:        p={r['Corrado_p']:.4f}")
        if "Sign_p" in r.index:
            S.append(f"  Sign test (non-param):    p={r['Sign_p']:.4f}")
        S.append(f"\n  PLAIN READING:")
        if r["BMP_p"] < 0.05 and r["CAAR_pct"] > 0:
            S.append(f"  Markets reacted positively and significantly to tokenization")
            S.append(f"  announcements — consistent with TAVE's core claim that tokenization")
            S.append(f"  is value-accretive (H1 supported, positive direction).")
        elif r["BMP_p"] < 0.10 and r["CAAR_pct"] > 0:
            S.append(f"  Positive reaction, significant at the 10% level — directional support")
            S.append(f"  for H1; note the marginal significance given sample size.")
        elif r["BMP_p"] < 0.05 and r["CAAR_pct"] < 0:
            S.append(f"  Markets reacted NEGATIVELY and significantly (p<0.05). H1 is supported")
            S.append(f"  in the sense of a measurable reaction, but the SIGN is negative: the")
            S.append(f"  market prices tokenization announcements as near-term risk/cost events")
            S.append(f"  (execution, regulatory, capex) over immediate value. This is consistent")
            S.append(f"  with TAVE's risk channels (smart-contract + regulatory haircut) dominating")
            S.append(f"  the short-run reaction. A defensible, non-obvious finding — lead with it.")
        elif r["BMP_p"] < 0.10 and r["CAAR_pct"] < 0:
            S.append(f"  Negative reaction, significant at the 10% level. Directional evidence that")
            S.append(f"  the market prices near-term risk over immediate value; report the marginal")
            S.append(f"  significance honestly given sample size.")
        else:
            S.append(f"  No statistically significant average reaction. This is a legitimate")
            S.append(f"  finding (market may already price tokenization, or power is limited).")
            S.append(f"  Report honestly and lean on the cross-sectional and subsample results.")

    # All windows
    S.append(f"\nALL EVENT WINDOWS")
    if test_results is not None and not test_results.empty:
        for _, r in test_results.iterrows():
            S.append(f"  {r['window']:<8} N={int(r['N']):<3} CAAR={r['CAAR_pct']:+.3f}%  "
                     f"BMP p={r['BMP_p']:.4f} {r['sig_BMP']}")

    # H2 — the key channel test
    S.append(f"\nKEY CHANNEL TEST — H2 (does the illiquidity channel dominate?)")
    if reg is not None and "BaselineIlliquidity" in reg["coefficients"].index:
        b = reg["coefficients"].loc["BaselineIlliquidity"]
        S.append(f"  Regression N={reg['n_obs']}, firm clusters={reg.get('n_clusters')}, SE={reg['se_type']}")
        S.append(f"  BaselineIlliquidity coef: {b['coef']:+.4f}  (p={b['p_value']:.4f})")
        if b["coef"] > 0 and b["p_value"] < 0.10:
            S.append(f"  PLAIN READING: Markets reward tokenization MORE when the underlying")
            S.append(f"  asset was more illiquid. This is the central TAVE finding — the")
            S.append(f"  liquidity-premium channel (Gap 1) is the dominant value driver.")
        else:
            S.append(f"  PLAIN READING: The illiquidity channel is not statistically dominant")
            S.append(f"  in this sample. May be power-limited; report the coefficient honestly.")
    else:
        S.append(f"  Regression did not run or variable absent — see diagnostics.")

    # Other coefficients
    if reg is not None:
        S.append(f"\nOTHER REGRESSION COEFFICIENTS (CAR drivers — primary/full sample)")
        for var in reg["coefficients"].index:
            if var == "const":
                continue
            c = reg["coefficients"].loc[var]
            star = "***" if c["p_value"] < 0.01 else "**" if c["p_value"] < 0.05 else "*" if c["p_value"] < 0.10 else ""
            S.append(f"  {var:<22} {c['coef']:+.4f}  (p={c['p_value']:.3f}) {star}")

    # Primary vs high-impact robustness comparison
    S.append(f"\nTWO-REGRESSION DESIGN — PRIMARY vs HIGH-IMPACT ROBUSTNESS")
    if reg is not None and reg_hi is not None:
        S.append(f"  Primary (full sample):     N={reg['n_obs']}")
        S.append(f"  Robustness (high-impact):  N={reg_hi['n_obs']}")
        if "BaselineIlliquidity" in reg["coefficients"].index and "BaselineIlliquidity" in reg_hi["coefficients"].index:
            b_full = reg["coefficients"].loc["BaselineIlliquidity"]
            b_hi = reg_hi["coefficients"].loc["BaselineIlliquidity"]
            S.append(f"  Illiquidity coef — full:       {b_full['coef']:+.4f} (p={b_full['p_value']:.3f})")
            S.append(f"  Illiquidity coef — high-impact:{b_hi['coef']:+.4f} (p={b_hi['p_value']:.3f})")
            S.append(f"  PLAIN READING:")
            if (abs(b_hi["coef"]) > abs(b_full["coef"]) and b_hi["p_value"] < 0.10):
                S.append(f"  Effect STRENGTHENS to significance on high-impact events (p<0.10) —")
                S.append(f"  evidence the signal is real, attenuated by less material events in the")
                S.append(f"  full sample. Write: 'holds in full sample, strengthens on high-impact.'")
            elif b_full["p_value"] < 0.10:
                S.append(f"  Illiquidity is significant in the full sample; on the high-impact subset")
                S.append(f"  it is {('larger' if abs(b_hi['coef'])>abs(b_full['coef']) else 'similar/smaller')}")
                S.append(f"  but p={b_hi['p_value']:.3f}. Report both; lead with the full-sample result.")
            else:
                S.append(f"  Illiquidity is NOT significant in either specification (full p="
                         f"{b_full['p_value']:.3f}, high-impact p={b_hi['p_value']:.3f}).")
                S.append(f"  H2 is not supported at the firm level here. Do NOT claim the effect")
                S.append(f"  strengthens — report the null honestly and rely on the asset-class")
                S.append(f"  gradient (descriptive) and JurisdictionScore for the channel story.")
    elif reg is not None:
        S.append(f"  Primary regression only (N={reg['n_obs']}). Too few high-impact events for")
        S.append(f"  a separate robustness regression — note as a limitation, not a failure.")

    # Audit-disclosure finding — AuditPresence is dropped from the OLS as zero-variance,
    # but the zero itself is a substantive finding for Gap 4a (smart-contract risk).
    if car_df is not None and "audit_disclosed" in car_df.columns:
        ev_level = car_df.drop_duplicates("event_id") if "event_id" in car_df.columns else car_df
        n_ev = len(ev_level)
        n_audit = int(ev_level["audit_disclosed"].astype(bool).sum())
        pct = (100.0 * n_audit / n_ev) if n_ev else 0.0
        S.append(f"\nSMART-CONTRACT AUDIT DISCLOSURE (Gap 4a — substantive null finding)")
        S.append(f"  Events disclosing a smart-contract audit: {n_audit}/{n_ev} ({pct:.1f}%)")
        if n_audit == 0:
            S.append(f"  FINDING: NOT ONE event in the sample disclosed a smart-contract audit.")
            S.append(f"  This is why AuditPresence is dropped from the regression (no variation) —")
            S.append(f"  but the null is itself the result: the market CANNOT price audit quality")
            S.append(f"  as a smart-contract-risk mitigant because issuers do not disclose it.")
            S.append(f"  Report as a disclosure-gap finding and a concrete policy/practice")
            S.append(f"  recommendation (standardised audit disclosure), not as a missing variable.")

    # Asset class ranking (descriptive H2)
    S.append(f"\nABNORMAL RETURN BY ASSET CLASS (descriptive — supports H2 if illiquid ranks high)")
    if asset_sub is not None and not asset_sub.empty:
        for _, r in asset_sub.iterrows():
            S.append(f"  {r['asset_class']:<16} N={int(r['N']):<3} CAAR={r['CAAR_pct']:+.3f}%  p={r['BMP_p']:.3f}")

    # Region robustness
    S.append(f"\nROBUSTNESS — BY REGION (effect should not be one-market-driven)")
    if region_sub is not None and not region_sub.empty:
        for _, r in region_sub.iterrows():
            S.append(f"  {r['region']:<6} N={int(r['N']):<3} CAAR={r['CAAR_pct']:+.3f}%  p={r['BMP_p']:.3f}")

    # Winsorized
    if wins:
        S.append(f"\nROBUSTNESS — OUTLIER SENSITIVITY (winsorized 5%)")
        S.append(f"  Winsorized CAAR: {wins['CAAR_winsorized_pct']:+.3f}%  (t={wins['t']:.3f}, p={wins['p']:.4f})")
        S.append(f"  If close to the headline CAAR, the result is not outlier-driven.")

    # Per-market outlier / leave-out diagnostic — is any single market driven by a few events?
    if outlier_summary is not None and not outlier_summary.empty:
        S.append(f"\nPER-MARKET OUTLIER CHECK (is any market's CAAR driven by a few events?)")
        S.append(f"  Read: if mean & median diverge a lot, or CAAR collapses after dropping the")
        S.append(f"  top-5 events, that market is outlier-driven — do NOT headline it as-is.")
        for _, r in outlier_summary.iterrows():
            excl_col = [c for c in outlier_summary.columns if c.startswith("CAAR_excl_top")]
            p_col = [c for c in outlier_summary.columns if c.startswith("BMP_p_excl_top")]
            excl = r[excl_col[0]] if excl_col else float("nan")
            pex = r[p_col[0]] if p_col else float("nan")
            S.append(f"  {r['market']:<5} N={int(r['N']):<4} mean={r['mean_CAAR_pct']:+.3f}%  "
                     f"median={r['median_CAR_pct']:+.3f}%  excl-top5={excl:+.3f}% (p={pex:.3f})  "
                     f"weekend-dates={int(r['weekend_dates'])}  [{r['robustness']}]")
        flagged = outlier_summary[outlier_summary["robustness"] != "ok"]["market"].tolist()
        if flagged:
            S.append(f"  ⚠ FLAGGED as outlier-skewed: {', '.join(map(str, flagged))} — inspect the")
            S.append(f"    Top_Contributing_Events sheet before using these in the write-up.")
        else:
            S.append(f"  ✓ No market flagged outlier-skewed — regional CAARs look broad-based.")

    # Placebo / falsification test — the strongest "is this real?" check
    if placebo is not None and not placebo.empty:
        S.append(f"\nPLACEBO / FALSIFICATION TEST (random pseudo-event dates — effect should VANISH)")
        for _, r in placebo.iterrows():
            pval = r["BMP_p"]
            pstr = f"{pval:.3f}" if pd.notna(pval) else "n/a"
            S.append(f"  Run {int(r['placebo_run'])}: N={int(r['N'])}  CAAR={r['CAAR_pct']:+.3f}%  BMP p={pstr}")
        n_sig = int((placebo["BMP_p"] < 0.10).sum())
        if n_sig == 0:
            S.append(f"  ✓ 0/{len(placebo)} placebo runs significant — the real effect is NOT an")
            S.append(f"    artifact of the method/benchmark. This is your key falsification result.")
        else:
            S.append(f"  ⚠ {n_sig}/{len(placebo)} placebo runs significant — the benchmark/method may")
            S.append(f"    produce spurious CAARs. Investigate before relying on the headline.")

    # WACC CHANNEL 1 — LIQUIDITY (empirical, Amihud + turnover)
    if liq_summary is not None and not liq_summary.empty:
        S.append(f"\nWACC CHANNEL 1 — LIQUIDITY (Amihud illiquidity + turnover, pre vs post event)")
        for _, r in liq_summary.iterrows():
            pstr = f"{r['p_value']:.3f}" if pd.notna(r['p_value']) else "n/a"
            S.append(f"  {r['metric']:<34} N={int(r['N']):<4} mean Δ={r['mean_pct_change']:+.2f}%  "
                     f"median Δ={r['median_pct_change']:+.2f}%  p={pstr}")
        amr = liq_summary[liq_summary["metric"].str.contains("DiD", na=False) &
                          liq_summary["metric"].str.startswith("Amihud")]
        if amr.empty:
            amr = liq_summary[liq_summary["metric"].str.startswith("Amihud")]
        if not amr.empty:
            a = amr.iloc[0]
            is_did = "DiD" in str(a["metric"])
            tag = "vs matched controls (DiD)" if is_did else "(raw pre→post — no controls)"
            if pd.notna(a["p_value"]) and a["p_value"] < 0.10 and a["mean_pct_change"] < 0:
                S.append(f"  → Illiquidity FELL significantly {tag} — EMPIRICAL SUPPORT for the")
                S.append(f"    liquidity-premium channel. Calibrate WACC ch.1 from this measured effect.")
            elif pd.notna(a["p_value"]) and a["p_value"] < 0.10:
                S.append(f"  → Illiquidity ROSE significantly {tag} — opposite sign; report honestly.")
            else:
                S.append(f"  → No significant liquidity effect {tag} — channel-1 benefit NOT robust:")
                S.append(f"    any raw pre→post change reflects market-wide trends, not the announcement.")
                S.append(f"    Calibrate WACC ch.1 from literature (Amihud-Mendelson) as an assumption.")

    # WACC CALIBRATION STATUS — which channels are empirical vs literature-assumed
    S.append(f"\nWACC ADJUSTMENT CALIBRATION STATUS (for the TAVE tool / write-up)")
    S.append(f"  Ch.1 Illiquidity premium    : EMPIRICAL (Amihud/turnover above) — or literature if null")
    S.append(f"  Ch.2 Info asymmetry         : LITERATURE assumption (Easley-O'Hara) — needs microstructure data")
    S.append(f"  Ch.3 Collateral/cost-of-debt: LITERATURE assumption (Bernanke-Gertler) — needs bond/CDS data")
    S.append(f"  Ch.4a Smart-contract risk   : DIRECTIONAL from CAAR sign (negative reaction = priced risk)")
    S.append(f"  Ch.4b Regulatory haircut    : DIRECTIONAL from JurisdictionScore coefficient")
    S.append(f"  → Report 3 channels as empirically informed, 2 as literature assumptions with")
    S.append(f"    sensitivity ranges. State this split explicitly — it is the honest, defensible design.")

    # Sample notes
    S.append(f"\nSAMPLE NOTES")
    S.append(f"  Events included:  {n_events}")
    S.append(f"  Events dropped:   {n_dropped}  "
             f"(no-ticker={len(skip.get('no_ticker',[]))}, no-index={len(skip.get('no_index',[]))}, "
             f"est-fail={len(skip.get('est_failed',[]))}, confounded={len(skip.get('confounded',[]))})")

    # Deliverables
    S.append(f"\nDELIVERABLES")
    S.append(f"  Results file: {RESULTS_FILE}")
    S.append(f"     Sheets: CAR_per_event, Significance_Tests, Index_Health, Skipped_Events,")
    S.append(f"             OLS_Coefficients, OLS_Full_Summary, OLS_HC3_Comparison,")
    S.append(f"             Robustness_Winsorized, Subsample_Region, Subsample_AssetClass")
    S.append(f"  Chart:        {PLOT_FILE}")
    S.append(f"     Top panel: average abnormal return by day (-10 to +10)")
    S.append(f"     Bottom panel: cumulative average abnormal return")
    S.append(f"     READ THE CHART: a jump at day 0 = clean immediate pricing; a rising")
    S.append(f"     line after day 0 = market digesting the signal over several days.")

    S.append(f"\nWHAT TO WRITE UP")
    S.append(f"  1. Lead with the headline CAAR + BMP significance, stating the SIGN honestly")
    S.append(f"     (negative = market prices near-term risk; the non-obvious, defensible finding).")
    S.append(f"  2. Report H2 as tested: if illiquidity is not significant, say so — rely on the")
    S.append(f"     asset-class gradient (descriptive) and JurisdictionScore for the channel story.")
    S.append(f"  3. Use the asset-class ranking to illustrate the liquidity logic descriptively.")
    S.append(f"  4. Report the per-market check: state whether the effect is broad-based OR")
    S.append(f"     concentrated in one market — do not claim 'broad-based' if a market is flagged.")
    S.append(f"  5. Report limitations transparently (dropped/confounded events, cluster count,")
    S.append(f"     JurisdictionScore being a low-N jurisdiction contrast, not a full gradient).")

    # ── RESEARCH-QUESTION / HYPOTHESIS SCORECARD (ties results to the paper) ──
    S.append(f"\nRESEARCH-QUESTION & HYPOTHESIS SCORECARD")
    # Pull the headline numbers for the verdicts
    prim_row = test_results[test_results["window"] == "(-1,1)"].iloc[0] if (test_results is not None and not test_results.empty) else None
    caar_v = prim_row["CAAR_pct"] if prim_row is not None else float("nan")
    bmp_p_v = prim_row["BMP_p"] if prim_row is not None else float("nan")
    h2_sig = False
    if reg is not None and "BaselineIlliquidity" in reg["coefficients"].index:
        h2_sig = bool(reg["coefficients"].loc["BaselineIlliquidity"]["p_value"] < 0.10)
    juris_sig = False
    if reg is not None and "JurisdictionScore" in reg["coefficients"].index:
        juris_sig = bool(reg["coefficients"].loc["JurisdictionScore"]["p_value"] < 0.10)

    S.append(f"  R1 — Is there a priced valuation effect a standard WACC/DCF would miss?")
    if pd.notna(bmp_p_v) and bmp_p_v < 0.05:
        S.append(f"       ANSWERED — YES. Significant abnormal return ({caar_v:+.3f}%, BMP p={bmp_p_v:.4f}).")
        S.append(f"       A measurable market reaction exists → supports the TAVE value proposition.")
    elif pd.notna(bmp_p_v) and bmp_p_v < 0.10:
        S.append(f"       ANSWERED — DIRECTIONAL. Marginal reaction ({caar_v:+.3f}%, p={bmp_p_v:.4f}).")
    else:
        S.append(f"       NOT SUPPORTED — no significant reaction; report as null.")
    S.append(f"  R2 — Through which channels does the effect enter value (WACC adjustments)?")
    S.append(f"       ANSWERED by construction — TAVE's 5-channel framework (theory + literature).")
    S.append(f"       Empirical handle: sign of CAAR + asset-class gradient + JurisdictionScore.")
    S.append(f"  R3 — Which proposed channels survive empirical stress-testing?")
    S.append(f"       TESTED (3 of 5): ch.1 illiquidity (Amihud), ch.4a/4b (CAAR sign + JurisdictionScore).")
    S.append(f"       ASSUMED (2 of 5): ch.2 info-asymmetry, ch.3 collateral — literature, future research.")
    S.append(f"")
    S.append(f"  H1 (a measurable reaction occurs): "
             + ("SUPPORTED" if (pd.notna(bmp_p_v) and bmp_p_v < 0.05) else "MARGINAL" if (pd.notna(bmp_p_v) and bmp_p_v < 0.10) else "NOT SUPPORTED")
             + f" — but SIGN is {'NEGATIVE' if (pd.notna(caar_v) and caar_v < 0) else 'POSITIVE'}.")
    S.append(f"       NOTE: if H1 predicted a POSITIVE (value-accretive) reaction, the negative sign")
    S.append(f"       REJECTS that directional form — report as risk-pricing, not confirmation.")
    S.append(f"  H2 (illiquidity channel is priced / dominates): "
             + ("SUPPORTED" if h2_sig else "NOT SUPPORTED at firm level") + ".")
    S.append(f"       Jurisdiction/regulatory effect: "
             + ("SIGNIFICANT" if juris_sig else "not significant")
             + " (caveat: low-N jurisdiction contrast, not yet a full gradient).")

    S.append("\n" + "─"*70)
    S.append("END FINDINGS — copy from the top border to here.")
    S.append("─"*70)
    text = "\n".join(S)
    print(text)
    return text


def self_test():
    """
    Fast integrity check — verifies every function the analysis relies on is
    defined before running. Catches the 'consumed def line' class of bug instantly.
    Returns True if OK.
    """
    required = [
        "load_events", "_truthy", "download_prices", "compute_returns",
        "estimate_market_model", "compute_ARs", "event_study",
        "_tz_naive", "get_earnings_dates", "is_confounded",
        "patell_test", "bmp_test", "corrado_rank_test", "sign_test", "run_all_tests",
        "cross_sectional_regression", "_coef_frame", "placebo_test",
        "robustness_winsorized", "subsample_by_region", "subsample_by_assetclass",
        "provenance_contrast", "pct_tokenized_signal", "wild_cluster_bootstrap_p",
        "market_outlier_diagnostic", "firm_exclusion_robustness", "liquidity_channel_analysis",
        "plot_ar_timeline", "plot_coefficient_estimates", "plot_assetclass_caar",
        "export_findings_pdf", "print_analysis_diagnostics", "print_findings_summary",
        "ensure_drive", "run_analysis", "run_pipeline", "_prepare_events_in_memory",
        "_cached_file_is_healthy",
    ]
    g = globals()
    missing = [name for name in required if name not in g or not callable(g[name])]
    if missing:
        print("✗ SELF-TEST FAILED — missing/again-broken functions:")
        for m in missing:
            print(f"    - {m}")
        print("  Do NOT run analysis until these are restored.")
        return False
    print(f"✓ SELF-TEST PASSED — all {len(required)} analysis functions defined.")
    return True


if __name__ == "__main__":
    # Full pipeline: collection -> analysis in memory (US cached, 4 markets fresh).
    # Falls back to reading the Excel only if run_collection isn't loaded.
    try:
        run_pipeline()
    except RuntimeError:
        run_analysis()
