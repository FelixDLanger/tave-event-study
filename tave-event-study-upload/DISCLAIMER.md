# Disclaimer

## Not financial, investment, legal, or tax advice

This repository contains **academic research software**. Nothing in this
repository, the accompanying paper, or the valuation tool constitutes
investment, financial, legal, accounting, or tax advice, an offer or
solicitation to buy or sell any security or digital asset, or a
recommendation regarding any tokenization program.

The author is not a registered investment adviser, broker-dealer, or
financial analyst in any jurisdiction, and is not acting in any such
capacity. Users must perform their own due diligence and obtain advice
from appropriately licensed professionals before making any decision.

## What the model is, and is not

The Tokenized Asset Valuation Engine (TAVE) is a **decomposition framework**,
not a predictive model. It does not forecast asset prices, token prices, or
returns. Its output is a structured estimate of how tokenization may affect
components of a firm's weighted average cost of capital under **user-supplied
assumptions**, and it is only as reliable as those assumptions.

Channels carry **different evidentiary weight**, and the tool labels each one:

- **Empirical** - an effect measured in the underlying study.
- **Directional** - a sign established by evidence, magnitude not established.
- **Literature-assumed** - a magnitude imported from prior work, untested here,
  reported with a sensitivity range.

A calibrated assumption must never be read as an estimated effect.

## Known limitations

- The event study measures announcement-window reactions, not long-horizon
  programme value.
- Cross-sectional inference rests on a small number of firm clusters; the
  wild cluster bootstrap is the governing inference, not analytic
  standard errors.
- The liquidity result is control-adjusted but not randomized: firms
  self-select into tokenization.
- The framework operates on the discount rate; cash-flow effects of
  tokenization are outside its scope.
- The practitioner-perception layer (N = 12) is interpreted qualitatively
  and supports no confirmatory statistical claim.

## Data

Event data is collected from public regulatory disclosure sources. Survey
data is anonymous and aggregated; no personal identifiers were collected
or retained. No proprietary or non-public data is included in this
repository.

## Warranty

This software is provided "as is", without warranty of any kind, express or
implied, in accordance with the AGPL-3.0 license. The author accepts no
liability for any loss arising from use of this software or its output.
