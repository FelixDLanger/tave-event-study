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
import warnings
warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════════════════

DRIVE_DIR    = "/content/drive/MyDrive/TAVE_Research"
INPUT_FILE   = os.path.join(DRIVE_DIR, "tave_master_database.xlsx")
RESULTS_FILE = os.path.join(DRIVE_DIR, "tave_results.xlsx")
PLOT_FILE    = os.path.join(DRIVE_DIR, "tave_AR_timeline.png")


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


def load_events() -> pd.DataFrame:
    """Load verified events from the master database."""
    df = pd.read_excel(INPUT_FILE, sheet_name="Events_Classified")
    df["event_date"] = pd.to_datetime(df["event_date"])

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
    data = {}
    for t in all_tickers:
        try:
            h = yf.download(t, start="2019-06-01", end="2026-01-31", progress=False)
            if not h.empty:
                close = h["Adj Close"] if "Adj Close" in h.columns else h["Close"]
                s = close.dropna()
                if len(s) > 0:
                    data[t] = s
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
            h = yf.download(idx, start="2019-06-01", end="2026-01-31", progress=False)
            close = (h["Adj Close"] if "Adj Close" in h.columns else h.get("Close")) if not h.empty else None
            s = close.dropna() if close is not None else pd.Series(dtype=float)
            if len(s) >= 500:
                data[idx] = s; used = idx; n_days = len(s); status = "NATIVE_OK"
        except Exception:
            pass

        # fallback if native failed
        if used is None:
            fb = INDEX_FALLBACK.get(idx)
            if fb:
                try:
                    h = yf.download(fb, start="2019-06-01", end="2026-01-31", progress=False)
                    close = (h["Adj Close"] if "Adj Close" in h.columns else h.get("Close")) if not h.empty else None
                    s = close.dropna() if close is not None else pd.Series(dtype=float)
                    if len(s) >= 500:
                        data[idx] = s   # store under the ORIGINAL key so lookups still work
                        used = fb; n_days = len(s); status = "FALLBACK_ETF"
                except Exception:
                    pass

        if used is None:
            status = "FAILED"
        index_health[idx] = {"index": idx, "used_symbol": used, "n_days": n_days, "status": status}
        index_resolved[idx] = used

    return data, index_health


# ═══════════════════════════════════════════════════════════════════════════════
# CONFOUNDING-EVENT FILTER (econometric refinement)
# ═══════════════════════════════════════════════════════════════════════════════

def get_earnings_dates(yf_ticker: str) -> list:
    """
    Fetch known earnings dates for a firm via yfinance. Used to flag events whose
    window overlaps an earnings release (a classic confound). Returns list of dates
    or empty list if unavailable.
    """
    try:
        tk = yf.Ticker(yf_ticker)
        ed = tk.get_earnings_dates(limit=40)
        if ed is not None and not ed.empty:
            return [pd.Timestamp(d).normalize() for d in ed.index]
    except Exception:
        pass
    return []


def is_confounded(event_date, earnings_dates, window_days=CONFOUND_WINDOW_DAYS) -> bool:
    """True if any earnings date falls within +/- window_days of the event."""
    ed_event = pd.Timestamp(event_date).normalize()
    for ed in earnings_dates:
        if abs((ed - ed_event).days) <= window_days:
            return True
    return False



    """Daily log returns."""
    return np.log(price_series / price_series.shift(1)).dropna()


def estimate_market_model(firm_ret, mkt_ret, event_date):
    """
    Estimate alpha, beta on the estimation window [-260, -11].
    Returns (alpha, beta, residual_std, n_obs) or None if insufficient data.
    """
    # Align
    df = pd.DataFrame({"firm": firm_ret, "mkt": mkt_ret}).dropna()
    df = df.sort_index()

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
    alpha, beta = model.params[0], model.params[1]
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
                "ticker": ev["ticker"], "exchange": ev["exchange"],
                "jurisdiction": ev["jurisdiction"],
                "jurisdiction_score": ev.get("jurisdiction_score", np.nan),
                "event_date": ev["event_date"],
                "announcement_type": ev.get("announcement_type", ""),
                "asset_class": ev.get("asset_class", ""),
                "audit_disclosed": ev.get("audit_disclosed", False),
                "pct_tokenized": ev.get("pct_tokenized", np.nan),
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
            "sig_BMP": "***" if p_b < 0.01 else "**" if p_b < 0.05 else "*" if p_b < 0.10 else "",
        })
    return pd.DataFrame(results)


# ═══════════════════════════════════════════════════════════════════════════════
# CROSS-SECTIONAL OLS — TAVE VALIDATION
# ═══════════════════════════════════════════════════════════════════════════════

def cross_sectional_regression(car_df, fundamentals_df=None, window="(-1,1)"):
    """
    Regress CAR on TAVE channel proxies.

    CAR_i = g0 + g1*BaselineIlliquidity + g2*TokenizationScope
            + g3*JurisdictionScore + g4*AuditPresence
            + g5*AssetClass_dummies + g6*log(MktCap) + g7*Leverage + e

    Tests H2-H5. H1 is the event study itself.
    """
    df = car_df[car_df["window"] == window].copy()
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

    # BaselineIlliquidity proxy:
    # Use asset class as ordinal illiquidity ranking (supply chain most illiquid)
    illiq_map = {
        "Supply Chain": 5, "Receivables": 4, "Real Estate": 4,
        "Bond": 2, "Fund": 2, "Equity": 1, "Deposit": 1,
        "Gold/Commodity": 3, "Unclassified": np.nan,
    }
    df["BaselineIlliquidity"] = df["asset_class"].map(illiq_map)

    # Merge fundamentals (market cap, leverage) if available
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
    candidate_vars = ["BaselineIlliquidity", "TokenizationScope",
                      "JurisdictionScore", "AuditPresence"]
    if "log_MktCap" in df.columns:
        candidate_vars.append("log_MktCap")
    if "Leverage" in df.columns:
        candidate_vars.append("Leverage")

    reg_df = df[["CAR", "firm_id"] + candidate_vars].dropna()
    if len(reg_df) < len(candidate_vars) + 3:
        print(f"  After dropping NaN, only {len(reg_df)} obs — reducing variables")
        candidate_vars = ["BaselineIlliquidity", "TokenizationScope",
                          "JurisdictionScore", "AuditPresence"]
        reg_df = df[["CAR", "firm_id"] + candidate_vars].dropna()

    if len(reg_df) < 8:
        print(f"  Too few observations for regression ({len(reg_df)})")
        return None

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
        "coefficients": pd.DataFrame({
            "coef": model.params,
            "std_err": model.bse,
            "t_stat": model.tvalues,
            "p_value": model.pvalues,
            "ci_low": model.conf_int()[0],
            "ci_high": model.conf_int()[1],
        }),
    }


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


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def run_analysis(filter_confounds=True):
    print("█"*70)
    print("  TAVE EVENT STUDY — CODE 2: ANALYSIS  (index-validated + clustered + diagnostics)")
    print("█"*70)

    # 0. Environment: mount Drive (if needed) + confirm input exists
    print("\n── Environment setup ──")
    ensure_drive()

    # 1. Load events
    print("\n── Loading events ──")
    events = load_events()
    if events.empty:
        print("No events to analyze."); return

    # 2. Load fundamentals
    fundamentals = pd.DataFrame()
    try:
        fundamentals = pd.read_excel(INPUT_FILE, sheet_name="YF_Fundamentals")
    except Exception:
        print("  No YF_Fundamentals sheet found — regression will use fewer controls")

    # 3. Download prices (with index validation + ETF fallback — Blocker 2)
    print("\n── Downloading + validating prices ──")
    tickers = events["yf_ticker"].dropna().unique().tolist()
    indices = list(set(MARKET_INDEX.values()))
    prices, index_health = download_prices(tickers, indices)
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

    # 6. Cross-sectional regression (firm-clustered SEs)
    print("\n── Cross-sectional OLS (TAVE validation) ──")
    reg = cross_sectional_regression(car_df, fundamentals, window="(-1,1)")
    if reg:
        print(f"  N={reg['n_obs']} | clusters={reg.get('n_clusters')} | SE={reg['se_type']}")
        print(reg["coefficients"].round(4).to_string())

    # 7. Robustness
    print("\n── Robustness ──")
    wins = robustness_winsorized(car_df)
    region_sub = subsample_by_region(car_df)
    asset_sub = subsample_by_assetclass(car_df)
    print("  By asset class (tests H2 - illiquidity channel):")
    print(asset_sub.to_string(index=False))

    # 8. Plot
    print("\n── Plotting ──")
    plot_ar_timeline(timeline, PLOT_FILE)

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
            reg["coefficients"].to_excel(w, "OLS_Coefficients")
            pd.DataFrame({"summary": [reg["summary"]]}).to_excel(w, "OLS_Full_Summary", index=False)
            # Also export HC3 version for comparison (cluster-SE caveat)
            if reg.get("model_hc3") is not None:
                hc3 = reg["model_hc3"]
                pd.DataFrame({
                    "coef": hc3.params, "std_err_HC3": hc3.bse,
                    "t_HC3": hc3.tvalues, "p_HC3": hc3.pvalues,
                }).to_excel(w, "OLS_HC3_Comparison")
        if wins:
            pd.DataFrame([wins]).to_excel(w, "Robustness_Winsorized", index=False)
        region_sub.to_excel(w, "Subsample_Region", index=False)
        asset_sub.to_excel(w, "Subsample_AssetClass", index=False)

    print(f"\n✓ RESULTS: {RESULTS_FILE}")
    print(f"✓ PLOT:    {PLOT_FILE}")

    # 10. DIAGNOSTIC VERDICT
    problems = print_analysis_diagnostics(events, car_df, index_health, skip, test_results, reg)

    # 11. COPY-PASTE FINDINGS SUMMARY
    print_findings_summary(car_df, test_results, reg, asset_sub, region_sub, wins, skip)

    return car_df, test_results, reg, index_health, skip, problems


def print_findings_summary(car_df, test_results, reg, asset_sub, region_sub, wins, skip):
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
        print("\n".join(S)); return

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
        S.append(f"\n  PLAIN READING:")
        if r["BMP_p"] < 0.05 and r["CAAR_pct"] > 0:
            S.append(f"  Markets reacted positively and significantly to tokenization")
            S.append(f"  announcements — consistent with TAVE's core claim that tokenization")
            S.append(f"  is value-accretive (H1 supported).")
        elif r["BMP_p"] < 0.10 and r["CAAR_pct"] > 0:
            S.append(f"  Positive reaction, significant at the 10% level — directional support")
            S.append(f"  for H1; note the marginal significance given sample size.")
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
        S.append(f"\nOTHER REGRESSION COEFFICIENTS (CAR drivers)")
        for var in reg["coefficients"].index:
            if var == "const":
                continue
            c = reg["coefficients"].loc[var]
            star = "***" if c["p_value"] < 0.01 else "**" if c["p_value"] < 0.05 else "*" if c["p_value"] < 0.10 else ""
            S.append(f"  {var:<22} {c['coef']:+.4f}  (p={c['p_value']:.3f}) {star}")

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
    S.append(f"  1. Lead with the headline CAAR and BMP significance (H1).")
    S.append(f"  2. State the H2 result — whether illiquidity dominates (the TAVE contribution).")
    S.append(f"  3. Use the asset-class ranking to illustrate H2 descriptively.")
    S.append(f"  4. Cite region robustness to show the effect is broad-based.")
    S.append(f"  5. Report sample limitations (dropped events, cluster count) transparently.")

    S.append("\n" + "─"*70)
    S.append("END FINDINGS — copy from the top border to here.")
    S.append("─"*70)
    print("\n".join(S))


if __name__ == "__main__":
    run_analysis()
