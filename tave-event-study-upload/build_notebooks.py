#!/usr/bin/env python3
"""Build three Colab-ready .ipynb notebooks for the TAVE pipeline."""
import json, os

OUT = "/mnt/user-data/outputs"

def md(*lines):
    return {"cell_type": "markdown", "metadata": {}, "source": [l if l.endswith("\n") else l+"\n" for l in lines]}

def code(src):
    if isinstance(src, str):
        src = src.splitlines(keepends=True)
    return {"cell_type": "code", "metadata": {}, "execution_count": None, "outputs": [], "source": src}

def notebook(cells):
    return {
        "cells": cells,
        "metadata": {
            "colab": {"provenance": [], "toc_visible": True},
            "kernelspec": {"name": "python3", "display_name": "Python 3"},
            "language_info": {"name": "python"}
        },
        "nbformat": 4, "nbformat_minor": 0
    }

def write(nb, name):
    path = os.path.join(OUT, name)
    with open(path, "w") as f:
        json.dump(nb, f, indent=1)
    print(f"  wrote {name} ({len(nb['cells'])} cells)")

# Read the source files
def _strip_main_guard(src):
    """Remove the trailing `if __name__ == "__main__":` block before embedding a
    module into a notebook cell. In Colab a cell's __name__ IS "__main__", so an
    auto-run guard would fire a full collection/analysis (and a SCRAPE) merely on
    loading the cell. Notebook cells must be pure libraries — define functions only;
    the explicit run cell is the single entry point. The .py files keep their guards
    so they still work when executed as standalone scripts."""
    marker = 'if __name__ == "__main__":'
    idx = src.find(marker)
    if idx == -1:
        return src
    # keep everything before the guard; add a one-line note in its place
    return src[:idx].rstrip() + (
        "\n\n# (Module-load auto-run guard intentionally omitted in the notebook —\n"
        "#  loading this cell only DEFINES functions; use the explicit run cell.)\n")

with open(os.path.join(OUT, "tave_code1_collection.py")) as f:
    code1 = _strip_main_guard(f.read())
with open(os.path.join(OUT, "tave_code2_analysis.py")) as f:
    code2 = _strip_main_guard(f.read())

# Shared bootstrap cells
boot_install = code(
"# Install all dependencies (run once per Colab session)\n"
"!pip install -q requests pandas numpy scipy statsmodels yfinance simfin \\\n"
"    tqdm beautifulsoup4 lxml openpyxl matplotlib"
)

boot_drive = code(
"# Mount Google Drive (the scripts also self-mount, but doing it here is clean)\n"
"from google.colab import drive\n"
"drive.mount('/content/drive')"
)

boot_secrets = code(
"# Secrets check — add these via the KEY ICON in the left sidebar:\n"
"#   SIMFIN_API_KEY (or SIMFIN_KEY) = your SimFin API key  [optional — yfinance covers fundamentals]\n"
"#   SEC_EMAIL                      = your email (SEC EDGAR fair-access header)\n"
"# The code reads them automatically through Colab's secret manager.\n"
"from google.colab import userdata\n"
"def _check(names):\n"
"    for n in names:\n"
"        try:\n"
"            if userdata.get(n):\n"
"                return f'{n}: set ✓'\n"
"        except Exception:\n"
"            pass\n"
"    return f'{names[0]}: MISSING ✗  (add via key icon, then allow notebook access)'\n"
"print(_check(['SIMFIN_API_KEY', 'SIMFIN_KEY']))\n"
"print(_check(['SEC_EMAIL']))"
)

# ──────────────────────────────────────────────────────────────────
# NOTEBOOK 1 — COLLECTION
# ──────────────────────────────────────────────────────────────────
nb1 = notebook([
    md("# TAVE — Code 1: Event Collection & Classification",
       "",
       "Collects tokenization announcements across **US (SEC EDGAR), Singapore (SGX), and "
       "Switzerland (SIX)** — three diversified jurisdictions across the Americas, Asia, and Europe. "
       "(LSE, HKEX, and TSE are deferred to future research: their data sits behind commercial "
       "licensing or fragile endpoints.)",
       "",
       "Auto-classifies, de-duplicates, scores each event for impact, auto-verifies, pulls "
       "financials, and writes one Excel file to `Drive/TAVE_Research/tave_master_database.xlsx`.",
       "",
       "**Before running:** add two Colab secrets via the key icon (left sidebar): "
       "`SIMFIN_API_KEY` (optional) and `SEC_EMAIL`.",
       "Run the cells top to bottom."),
    md("## 1. Setup"),
    boot_install,
    boot_drive,
    boot_secrets,
    md("## 2. Load the collection code",
       "",
       "This cell contains the full Code 1. (If you prefer, replace it with a `git clone` of your repo.)"),
    code(code1),
    md("## 2b. Integrity self-test (5-second check)",
       "",
       "Verifies every pipeline function is defined before a long run — catches any "
       "code-paste or edit error instantly. If this fails, do not proceed."),
    code("assert self_test(), 'Self-test failed — see missing functions above.'"),
    md("## 3. Run — US only first (validation)",
       "",
       "Fast pass to confirm EDGAR and the pipeline work. Read the **COPY-PASTE SUMMARY** at the bottom of the output."),
    code("# Validation run — US only.\n"
         "run_collection(markets=('US',), pull_financials=False, check_prices=False)"),
    md("## 4. Run — full collection (US + SGX + SIX)",
       "",
       "Once US looks healthy, run all three markets: collection + classification + dedup + "
       "impact scoring + financials + price coverage.",
       "",
       "**Caching:** US is reused from the saved file if already collected (saves ~15 min). "
       "To force a fresh re-scrape of a market, use `refresh=('SGX',)`. To ignore the cache "
       "entirely, use `use_cache=False`."),
    code("run_collection()                       # reuses cached US, scrapes SGX + SIX\n"
         "# run_collection(refresh=('SGX',))     # force re-scrape SGX only\n"
         "# run_collection(use_cache=False)      # re-scrape everything from scratch"),
    md("## 5. Verification (short, mostly automated)",
       "",
       "Code 1 already auto-verified clear events and auto-rejected noise. Only **AMBIGUOUS** rows need you.",
       "",
       "1. Open `Drive/TAVE_Research/tave_master_database.xlsx`",
       "2. Go to the **Manual_Review** sheet",
       "3. For each genuine tokenization event, set `verified = TRUE` on its row in **Events_Classified** "
       "(you can type TRUE, yes, x, or 1 — all accepted) and fix `asset_class` / `announcement_type` if wrong",
       "4. Save",
       "",
       "Then move to the **Analysis** notebook.")
])
write(nb1, "TAVE_1_Collection.ipynb")

# ──────────────────────────────────────────────────────────────────
# NOTEBOOK 2 — ANALYSIS
# ──────────────────────────────────────────────────────────────────
nb2 = notebook([
    md("# TAVE — Code 2: Event Study & Regression",
       "",
       "Reads `tave_master_database.xlsx` (from Code 1), runs the market-model event study, three "
       "significance tests (BMP primary, Patell, Corrado), and **two cross-sectional regressions** "
       "with firm-clustered standard errors: a PRIMARY model on the full verified sample and a "
       "ROBUSTNESS model on the high-impact subset. Adds confound filtering and robustness checks, "
       "then writes `tave_results.xlsx` plus `tave_AR_timeline.png` to Drive.",
       "",
       "**Run Code 1 first** — this notebook needs the file it produces."),
    md("## 1. Setup"),
    boot_install,
    boot_drive,
    md("## 2. Load the analysis code"),
    code(code2),
    md("## 2b. Integrity self-test (5-second check)",
       "",
       "Verifies every analysis function is defined before running. If this fails, do not proceed."),
    code("assert self_test(), 'Self-test failed — see missing functions above.'"),
    md("## 3. Run the analysis",
       "",
       "Prints the **COPY-PASTE FINDINGS** block at the end — paste that for interpretation."),
    code("run_analysis()"),
    md("## 4. (Optional) Disable the earnings-confound filter",
       "",
       "If the filter drops too many events on a thin sample, run this instead and report both versions."),
    code("# run_analysis(filter_confounds=False)")
])
write(nb2, "TAVE_2_Analysis.ipynb")

# ──────────────────────────────────────────────────────────────────
# NOTEBOOK 3 — COMBINED (both, end to end)
# ──────────────────────────────────────────────────────────────────
nb3 = notebook([
    md("# TAVE — Full Pipeline (Cache-First)",
       "",
       "End to end in one notebook. The single run cell is cache-first: if the master "
       "file holds a healthy 5-market sample it analyses that file DIRECTLY — no scrape, "
       "deterministic reruns. Loading the module cells only DEFINES functions (the "
       "auto-run guard is stripped), so nothing scrapes until the explicit run cell.",
       "",
       "**Secrets (key icon, left sidebar):** `SIMFIN_API_KEY` (optional) and `SEC_EMAIL`."),
    md("## 1. Setup"),
    boot_install,
    boot_drive,
    boot_secrets,
    md("## 2. Load Code 1 (collection — defines functions only; used as scrape fallback)"),
    code(code1),
    code("# Integrity self-test — all collection functions defined?\n"
         "assert self_test(), 'Self-test failed — see missing functions above.'"),
    md("## 3. Load Code 2 (analysis — defines functions only)"),
    code(code2),
    code("# Integrity self-test — all analysis functions defined?\n"
         "assert self_test(), 'Self-test failed — see missing functions above.'"),
    md("## 3b. OPTIONAL — Sample extension (additive, idempotent)",
       "",
       "Adds the verified extension firms (BNY Mellon, Mastercard, PayPal, Broadridge, "
       "WisdomTree, Hamilton Lane, Abrdn) to the master file. **Flag-guarded so 'Run all' "
       "is safe**: set `RUN_EXTENSION = True` once, run, then set it back. Idempotent — "
       "firms already collected are skipped, so a second run costs seconds."),
    code("RUN_EXTENSION = True    # ARMED - idempotent (collected firms skipped in seconds); set False after success" + chr(10) + "if RUN_EXTENSION:" + chr(10) + "    run_extension_collection()"),
    md("## 4. Run the pipeline (cache-first; scrape only as fallback)",
       "",
       "**The only run cell.** `prefer_cache=True`: a healthy master file is analysed "
       "directly with no scrape, so reruns are deterministic and your clean data is never "
       "overwritten. Use `prefer_cache=False` only when you deliberately want fresh data."),
    code("run_pipeline(\n"
         "    markets=('US', 'SGX', 'SIX', 'XETRA', 'LSE'),\n"
         "    refresh=('SGX', 'SIX', 'XETRA', 'LSE'),  # only used if the cache is unhealthy\n"
         "    prefer_cache=True,                       # healthy file -> analyse directly, no scrape\n"
         ")")
])
write(nb3, "TAVE_Full_Pipeline.ipynb")

print("\nDone. Three notebooks created.")
