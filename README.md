# TAVE — Tokenized Asset Valuation Engine

Replication code and open-source valuation tool for the paper:

> **Tokenization and the Cost of Capital: A Five-Channel Valuation Framework with Event-Study and Practitioner Evidence**
> Felix D. Langer — Independent Researcher; Doctoral Candidate (DBA), GlobalNxt University
> ORCID: [0009-0000-3205-8088](https://orcid.org/0009-0000-3205-8088)
> SSRN: *abstract ID to be added on posting*

---

## What this is

Enterprise asset tokenization is moving from pilot to production, but valuation
practice has no line item for the economic changes it produces. **TAVE** maps
tokenization onto the weighted average cost of capital through five named
channels, each anchored in established asset-pricing theory, and tests the
framework on the largest verified sample of enterprise tokenization events
assembled to date.

```
WACC_TAVE = WACC_base
            − Δ-ILP − Δ-IAP − Δ-Kd × (D/V)      ← benefit channels
            + SC_spread + Reg_spread             ← risk channels
```

![TAVE framework](TAVE_Framework_Figure1.png)

## Headline results

| Finding | Evidence |
|---|---|
| Markets price tokenization announcements **negatively** | CAAR −0.260%, BMP *p* = 0.002, N = 1,169; significant at 1% in all five windows |
| **No** cross-sectional covariate robustly differentiates the reaction | Jurisdiction effect collapses: wild cluster bootstrap *p* = 0.31; *p* = 0.85 ex-one-firm; absorbed by announcement-type controls |
| The **liquidity mechanism is real but unpriced** | Amihud price impact −2.70% vs matched non-announcing peers (*p* < 0.001); baseline illiquidity not priced cross-sectionally (*p* = 0.63) |
| Complete **audit-disclosure null** | 0 of 1,169 events disclosed an independent smart-contract audit |

A methodological note worth flagging: an apparently decisive jurisdiction
result (analytic *p* < 0.001, *t* = −9.0) proved to be a **few-cluster
artifact**, detected by this repository's own robustness battery. The code that
falsifies the finding ships alongside the code that produced it.

## Repository contents

| File | Purpose |
|---|---|
| `tave_code1_collection.py` | Event collection from primary disclosure systems (SEC EDGAR, SGXNet, SIX, XETRA, LSE), classification, de-duplication, canary verification, sample extension |
| `tave_code2_analysis.py` | Event study, significance tests, cross-sectional regressions, wild cluster bootstrap, matched-control DiD, placebo battery, exports |
| `build_notebooks.py` | Deterministic builder that regenerates the Colab notebooks from the two source modules |
| `tokenization-valuation.html` | The TAVE valuation tool — single-file, no build step, evidence-labeled parameters |
| `TAVE_Framework_Figure1.png` / `.svg` | Framework diagram (Figure 1 of the paper) |
| `CITATION.cff` | Machine-readable citation metadata |
| `DISCLAIMER.md` | Scope, limitations, and what the model is **not** |

## Reproducing the results

**Requirements.** Python 3.10+, `pandas`, `numpy`, `statsmodels`, `scipy`,
`yfinance`, `openpyxl`, `matplotlib`, `requests`, `beautifulsoup4`.
Designed to run in Google Colab; no local setup required.

**Credentials.** Two values are read from the environment and are **never**
stored in this repository:

| Variable | Required | Purpose |
|---|---|---|
| `SEC_EMAIL` | Yes | Identifying email for SEC EDGAR fair-access headers |
| `SIMFIN_API_KEY` | Optional | Fundamentals; `yfinance` covers the default path |

**Run.**

```bash
python build_notebooks.py          # regenerate the pipeline notebook
# then open TAVE_Full_Pipeline.ipynb and Run All
```

The pipeline is **cache-first**: a healthy master file is analyzed directly
with no re-scrape, so reruns are deterministic. Set `prefer_cache=False` only
to deliberately refresh source data. The master database is *not* committed —
it regenerates from public sources.

**Verification built in.** Every run performs a self-test of all defined
functions, canary-firm checks per market (known anchor events must be
present), index-health validation, and a read-back integrity check before
analysis.

## The valuation tool

`tokenization-valuation.html` runs in any browser with no server or build step.
Its design principle is that **evidence status travels with every parameter**:

- **Empirical** — effect measured in this study (the liquidity channel)
- **Directional** — sign established, magnitude not (smart-contract risk)
- **Literature-assumed** — magnitude imported with a sensitivity range, untested here (information, collateral, and — after the robustness battery — regulatory)

The tool deliberately produces a **decomposition, not a point estimate**,
because the evidence does not support one.

## Data sources

Event data derives entirely from public regulatory disclosure channels. Survey
data is anonymous and aggregated; the complete item-level response matrix
(12 × 24) is published in Appendix F of the paper. No personal identifiers were
collected or retained, and no proprietary or non-public data is included here.

## Citation

```bibtex
@article{langer2026tave,
  author = {Langer, Felix D.},
  title  = {Tokenization and the Cost of Capital: A Five-Channel Valuation
            Framework with Event-Study and Practitioner Evidence},
  year   = {2026},
  note   = {SSRN working paper},
  url    = {https://github.com/FelixDLanger/tave-event-study}
}
```

See `CITATION.cff` for machine-readable metadata.

## License

**AGPL-3.0-or-later.** Network use counts as distribution: if you deploy a
modified version as a service, you must publish your source. For commercial
licensing outside AGPL terms, contact the author.

## Disclaimer

This is academic research software. It is **not** investment, financial,
legal, or tax advice, and it does not forecast prices or returns. See
[`DISCLAIMER.md`](DISCLAIMER.md) for the full statement, including the known
limitations of the design.
