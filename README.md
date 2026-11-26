# JPMorgan Chase Bank Valuation and Risk Model

An independent, AI-assisted financial model of JPMorgan Chase & Co. (NYSE: JPM). Starting from reported data in SEC filings (FY2020 to 9M 2026), it:

- forecasts the bank's earnings and regulatory capital (CET1) through FY2031,
- values the equity three ways: FCFE DCF, residual income and justified price-to-tangible-book,
- stress-tests the result under Base, Bull and Bear scenarios, with sensitivity and tornado analysis.

The workbook is generated in Python with `openpyxl`. Every forecast and valuation cell is a live Excel formula, so any input can be changed in Excel and the model updates.

**Download the model:** [`model/JPM_Financial_Model.xlsx`](model/JPM_Financial_Model.xlsx)

## Key results (as of November 20, 2026)

| | Value per share | vs. $338.21 price |
| --- | --- | --- |
| FCFE DCF (excess-capital method) | $360 | +6% |
| Residual income | $349 | +3% |
| Justified price / tangible book | $344 | +2% |
| **Blended (50 / 25 / 25)** | **$353** | **+4%, Hold** |
| Bear scenario | $250 | -26% |
| Bull scenario | $422 | +25% |

**Stress test finding:** in a 2027 recession, JPM's CET1 ratio would fall from 14.4% to about 11.0% by 2031, below the ~11.5% regulatory minimum, if the bank kept paying out its base-case dividends and buybacks. Cutting buybacks is what keeps the bank above the floor.

**Largest value drivers (tornado analysis):** cost of equity, the credit-loss rate and the efficiency ratio. Balance-sheet growth matters least.

## Why a bank needs a different model

A bank's deposits and debt are its raw material, not its financing, so a standard free-cash-flow DCF with WACC does not apply. This model follows the approach bank analysts use:

- **Revenue** = average assets × net interest yield, plus fee income.
- **Costs** = revenue × efficiency ratio, plus credit-loss provisions on average loans.
- **Cash flow to shareholders (FCFE)** = net income minus the capital that must be retained to hold a target CET1 ratio as risk-weighted assets grow.
- **Discount rate** = cost of equity (CAPM), valuing equity directly rather than the enterprise.

Reported operating cash flow is shown only as a memo line: it swung from +$107B (2022) to -$148B (2025) with trading balances and carries no valuation signal.

## Workbook structure

| Tab | Contents |
| --- | --- |
| Cover | Live headline outputs, forecast snapshot, model map, methodology |
| Historicals | FY2020 to FY2025 and 9M 2026 reported data, with a net-income tie-out check |
| Assumptions | Scenario switch (cell C4), 4Q26 bridge, nine drivers × three scenarios × five years, shock cells |
| Forecast | Income statement FY2026E to FY2031E, EPS, ROE, ROTCE |
| Capital & Cash Flow | CET1 roll-forward, RWA, FCFE, dividends, buybacks, share count, tangible book value |
| Valuation | CAPM cost of equity, FCFE DCF, residual income, justified P/TBV, blended value |
| Scenarios & Sensitivity | Live sensitivity grids, scenario comparison, CET1 stress path, tornado chart, risk register |
| Sources | Origin of every hard-coded number |

Colour code: blue text = input, black = formula, green = link to another sheet, yellow fill = key lever.

## Repository layout

```
src/
  build_model.py     builds the formula-driven workbook
  run_scenarios.py   re-runs the model per scenario and per shock; writes the risk tables and charts
  recalc.py          recalculates a workbook with LibreOffice (headless)
model/
  JPM_Financial_Model.xlsx   the finished model
  layout.json                row map used by run_scenarios.py
  scenario_results.json      outputs of the latest scenario and tornado runs
```

## How to run

Requirements: Python 3.9+, `openpyxl`, and [LibreOffice](https://www.libreoffice.org/) (used only to recalculate formulas during the scenario runs).

```bash
pip install -r requirements.txt
python src/build_model.py      # writes model/JPM_Financial_Model.xlsx
python src/run_scenarios.py    # about 20 model re-runs; adds the risk-assessment tables
```

To try a different view, edit the inputs in `build_model.py` (or directly in the Excel file) and re-run both scripts. Tables 1 and 2 on the Scenarios & Sensitivity tab update live in Excel; Tables 3 to 5 refresh when `run_scenarios.py` is re-run.

## AI-assisted workflow

The model was built with Claude (Anthropic) as an analytical assistant:

1. **Data extraction:** reported figures were pulled from the SEC EDGAR XBRL API and JPM filings, then reconciled. Revenue equals net interest income plus noninterest revenue, and modelled net income ties to reported net income to the dollar for every year.
2. **Model generation:** the workbook is produced programmatically, with formulas rather than pasted values, so it stays auditable in Excel.
3. **Automated risk analysis:** scenario and tornado runs re-calculate the full model for each case instead of approximating.
4. **Review:** modelling choices were checked against results. For example, an early version penalised faster growth in the terminal value; the terminal FCFE was normalised so retained capital is consistent with terminal growth.

## Data sources

- SEC EDGAR XBRL company-concept API, CIK 0000019617 (FY2021 to FY2025 income statement, assets, deposits, equity, per-share data, buybacks, cash flow)
- JPMorgan Chase Annual Reports (2022, 2025): loans, common equity and CET1 ratios, including FY2020 balances
- Form 10-Q and earnings release for the quarter ended September 30, 2026
- Share price and 10-year Treasury yield as of November 16, 2026; beta from Yahoo Finance (5-year monthly)

Full citations are on the Sources tab of the workbook.

## Simplifications

- Firm-level forecast rather than a segment build.
- One net interest yield on total assets rather than separate asset yields and deposit costs.
- CET1 roll-forward holds AOCI, regulatory deductions and intangibles constant.
- The ~11.5% regulatory CET1 minimum and the 4Q26 estimates are approximations.

## Disclaimer

This is an independent educational project, not affiliated with or endorsed by JPMorgan Chase & Co. It is not investment advice.
