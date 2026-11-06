# Nifty 50 Monthly Portfolio Rebalancer

A standalone Python script that runs a monthly, rules-based rebalance of a Nifty 50 portfolio. It reads last month's holdings and the latest prices from an Excel workbook, decides which stocks to **ADD**, **RETAIN** or **REMOVE**, sizes the target positions, and writes an updated workbook that the next run picks up automatically.

> **Disclaimer:** This project is for educational and research purposes only. It is not financial advice, and it is not a recommendation to buy or sell any security. Backtest results and rule-based screens do not guarantee future returns. Check every trade yourself before placing it.

---

## Table of Contents

- [Features](#features)
- [How the Strategy Works](#how-the-strategy-works)
- [Repository Contents](#repository-contents)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Command-Line Options](#command-line-options)
- [Input: The Excel Workbook](#input-the-excel-workbook)
- [Output](#output)
- [Console Report](#console-report)
- [Monthly Workflow](#monthly-workflow)
- [Configuration and Defaults](#configuration-and-defaults)
- [How the Calculations Work](#how-the-calculations-work)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [Roadmap Ideas](#roadmap-ideas)
- [License](#license)

---

## Features

- **Relative-strength ranking** of all Nifty 50 constituents that trade above their 200-day moving average (DMA).
- **Concentrated portfolio** of up to 10 stocks, each sized at an equal weight.
- **Strict 200 DMA exit**: any holding at or below its 200 DMA is sold immediately.
- **Rank-retention buffer**: existing holdings are kept while they stay in the Top 15, which reduces monthly churn and short-term tax drag.
- **Market-breadth regime filter**: when at least 50% of the index is above its 200 DMA, the portfolio is 100% equity. Otherwise it moves to 50% equity and parks 50% in Nippon India ETF Liquid BeES (`LIQUIDBEES`).
- **Zero-manual-entry monthly loop**: holdings, share counts and cash are read from the workbook, and the updated state is written back to it.
- **Optional live data** via `yfinance` (`--live`), with automatic fallback to the workbook values.
- **Formatted Excel output** with colour-coded actions, live formulas, hyperlinks to Google Finance and NSE, and embedded bar charts.
- **Clear console report** showing the regime, breadth, and an action table for every position.

---

## How the Strategy Works

Each month the script applies four rules, in order.

| # | Rule | What it does |
|---|------|--------------|
| 1 | **Top-10 relative-strength concentration** | Ranks stocks above their 200 DMA by `% distance above 200 DMA = (LTP − 200 DMA) / 200 DMA`, then targets the top 10. |
| 2 | **Strict < 200 DMA exit** | Sells any holding where `LTP <= 200 DMA`, or that has left the Nifty 50 universe. |
| 3 | **Top-15 rank-retention buffer** | Keeps an existing holding if it is still above its 200 DMA and still ranks within the top 15. Holdings that are kept are capped at 10. Open slots are filled with the highest-ranked new candidates. |
| 4 | **Market-breadth regime filter** | If ≥ 50% of constituents are above their 200 DMA → **Bull**: 100% equity, 10% per stock. If < 50% → **Defensive**: 50% equity (5% per stock) and 50% in `LIQUIDBEES`. |

Position sizing: `target shares = floor(capital × per-stock weight ÷ LTP)`. Any leftover cash is reported as residual cash.

---

## Repository Contents

```text
.
├── nifty50_monthly_rebalancer.py   # The rebalancer script (single file)
├── Nifty50_Portfolio_Tracker.xlsx  # Tracker workbook: input on run 1, output thereafter
└── README.md                       # This file
```

---

## Requirements

- **Python 3.9 or newer** (developed and tested on Python 3.14)
- [`pandas`](https://pypi.org/project/pandas/)
- [`openpyxl`](https://pypi.org/project/openpyxl/)
- [`yfinance`](https://pypi.org/project/yfinance/) *(optional, only needed for `--live`)*

---

## Installation

1. Clone or download the repository:

   ```bash
   git clone https://github.com/<your-username>/<your-repo>.git
   cd <your-repo>
   ```

2. Install the dependencies:

   ```bash
   python -m pip install pandas openpyxl yfinance
   ```

   To skip live-data support, install only `pandas` and `openpyxl`.

3. Optionally, save the dependencies to a file for reproducible setups:

   ```text
   # requirements.txt
   pandas>=2.0
   openpyxl>=3.1
   yfinance>=0.2
   ```

   Then run `python -m pip install -r requirements.txt`.

---

## Quick Start

```bash
python nifty50_monthly_rebalancer.py
```

With no arguments, the script looks for `Nifty50_Portfolio_Tracker.xlsx` in the current folder, reads it, computes the rebalance, prints the action table, and overwrites the workbook with the new state.

> **Back up your workbook before running.** The script overwrites the file it reads.

---

## Command-Line Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `--excel PATH` | path | `Nifty50_Portfolio_Tracker.xlsx` | Workbook to read from and write to. |
| `--capital AMOUNT` | float | *(computed)* | Override total portfolio value in INR. |
| `--top-n N` | int | `10` | Number of stocks to hold. |
| `--keep-rank N` | int | `15` | Rank-retention buffer. Holdings ranked beyond this are sold. |
| `--breadth-cutoff X` | float | `0.50` | Share of constituents above 200 DMA needed for the Bull regime. |
| `--live` | flag | off | Refresh LTP and 200 DMA from Yahoo Finance (`.NS` tickers) before rebalancing. |
| `-h`, `--help` | flag | — | Show help and exit. |

**Examples**

```bash
# Use live prices and moving averages
python nifty50_monthly_rebalancer.py --live

# Run against a specific workbook with a fixed capital base
python nifty50_monthly_rebalancer.py --excel "D:\Investing\Tracker.xlsx" --capital 1500000

# Hold 8 stocks with a tighter buffer
python nifty50_monthly_rebalancer.py --top-n 8 --keep-rank 12
```

---

## Input: The Excel Workbook

The script reads two sheets. If the workbook does not exist, it starts from the built-in October 2026 baseline (see [Configuration and Defaults](#configuration-and-defaults)).

### `Current_Holdings` (portfolio state)

| Column | Meaning |
|--------|---------|
| `Symbol` | NSE ticker, for example `TECHM`. |
| `Held_Shares` | Number of shares currently held. |
| `Entry_Rebalance_LTP` | Price used at the last rebalance. |

Two special rows are recognised:

- `CASH_BALANCE`: uninvested cash in INR. The value is read from the `Entry_Rebalance_LTP` column.
- `LIQUIDBEES`: units of Liquid BeES held. Each unit is treated as ₹1,000 in the calculation.

### `Nifty50_Screener` (market data)

The script looks for a header row within the first 15 rows that contains a `Symbol` (or `Stock` / `Ticker`) column and an `LTP` column. It also reads columns containing `DMA` or `200`, and optionally `Company`/`Name` and `In_Portfolio`.

If the `Nifty50_Screener` sheet is missing, the script uses the first sheet that has these columns.

> The workbook produced by a previous run already has all of these columns, so the next run works with no edits.

---

## Output

Each run overwrites the workbook with three sheets:

1. **`Rebalance_Action_Sheet`**: one row per action (ADD, RETAIN, REMOVE) with LTP, 200 DMA, % vs DMA, previous and target shares, target weight and value, trade size, selection reason, and a source link. It includes a totals row and a bar chart of % distance vs 200 DMA. Most cells are live Excel formulas.
2. **`Current_Holdings`**: the active positions and `CASH_BALANCE`. This sheet is the input for the next run, so do not rename or restructure it.
3. **`Nifty50_Screener`**: all 50 constituents with LTP, 200 DMA, % vs DMA, above-DMA flag, RS rank, and portfolio membership. It includes an average row and a bar chart of the top 15.

Action colours: green = ADD, blue = RETAIN, red = REMOVE.

---

## Console Report

```text
====================================================================================
 NIFTY 50 MONTHLY REBALANCE REPORT — 2026-11-06 (READ FROM & SAVED TO: Nifty50_Portfolio_Tracker.xlsx)
====================================================================================
 Total Portfolio Capital : ₹1,631,397.87 (Invested: ₹1,618,270.80 | Residual Cash: ₹13,127.07)
 Nifty 50 Market Breadth : 32 / 50 stocks > 200 DMA (64.0%)
 Active Market Regime    : BULL REGIME (100% Equity)
 Monthly Actions         : 10 RETAIN | 0 ADD (Buy) | 0 REMOVE (Sell)
 Target Weightage        : 10.00% per equity stock | 0.00% Liquid ETF
...
```

Live-data status is printed before the report when `--live` is used.

---

## Monthly Workflow

1. **Before the market opens on the 1st**, back up the workbook (for example, copy it to `Archive\2026-12.xlsx`).
2. **Run the script** (add `--live` if you have internet access and want fresh data):
   ```bash
   python nifty50_monthly_rebalancer.py --live
   ```
3. **Review the action table** in the console and in `Rebalance_Action_Sheet`.
4. **Execute the trades** in your broker account, matching the `Trade Shares` column.
5. **Update `Current_Holdings`** if your actual fills differ from the targets. The script treats the workbook as the source of truth, so keep it in line with your real positions.
6. Repeat next month. The workbook carries the state forward.

---

## Configuration and Defaults

The following values are hard-coded at the top of the script.

- **`DEFAULT_NIFTY50_DATA`**: the Nifty 50 constituents with company name, LTP and 200 DMA. These are a **snapshot from November 2026** and are used only when the workbook has no screener data. Update this list when the index changes.
- **`INITIAL_OCT2026_HOLDINGS` / `INITIAL_OCT2026_CASH`**: the starting portfolio, used only when no holdings are found in the workbook.
- **Liquid BeES price**: fixed at ₹1,000 per unit.
- **Default capital** when there are no holdings and no `--capital` value: ₹10,00,000.
- **Strategy parameters**: `top_n=10`, `keep_rank=15`, `breadth_threshold=0.50`, `defensive_equity_wt=0.50`.

To change the strategy, edit the defaults in `run_rebalance_calculation()` or pass the matching command-line options.

---

## How the Calculations Work

- **Distance above 200 DMA**: `((LTP / DMA_200) − 1) × 100`.
- **Universe**: all stocks in the screener. Only stocks with `LTP > DMA_200` are ranked.
- **Breadth**: `number above 200 DMA ÷ total stocks`. Bull when this is ≥ the cutoff.
- **Portfolio value** (when `--capital` is not given): `Σ(previous shares × current LTP) + LIQUIDBEES units × 1,000 + cash`.
- **Retained holdings**: sorted by rank. If more than `top_n` qualify, the lowest-ranked are trimmed and sold.
- **Additions**: fill the open slots from the top-ranked stocks above their 200 DMA that are not already held.
- **Target shares**: `floor(total capital × weight ÷ LTP)`. `Trade Shares = target − previous`.
- **Liquid BeES**: `floor(total capital × liquid weight ÷ 1,000)` units.
- **Residual cash**: `total capital − Σ(target shares × LTP) − liquid value`.
- **`--live` data**: LTP is the most recent close, and the 200 DMA is the simple mean of the last 200 daily closes, using `period="1y"` from Yahoo Finance. Stocks with fewer than 200 closes keep the workbook values.

---

## Troubleshooting

| Symptom | Likely cause and fix |
|---------|----------------------|
| `ModuleNotFoundError: No module named 'openpyxl'` (or `pandas`) | Dependencies are not installed for this Python. Run `python -m pip install pandas openpyxl`. |
| `[INFO] Using LTP & 200 DMA values from the Excel workbook.` after `--live` | Live fetch failed (no internet, Yahoo rate limit, or `yfinance` not installed). The run continues with workbook values. |
| Wrong holdings or cash | Check `Current_Holdings` against your broker statement, then re-run. Restore from your backup if the file was already overwritten. |
| Screener not detected | Make sure the header row has `Symbol` and `LTP` (and a `200 DMA` column) within the first 15 rows. |
| `Permission denied` when saving | Close the workbook in Excel before running. |
| Chart or formula looks odd in Excel | Excel recalculates formulas on open. Press `F9` if values look stale. |

---

## Limitations

- **Not a backtester.** The script produces one month's rebalance from the current data. It does not simulate past performance.
- **No transaction costs, taxes, slippage or lot-size rules** are modelled. Share counts are rounded down, but brokerage, STT, stamp duty and exchange charges are not included.
- **Liquid BeES is priced at a fixed ₹1,000** in the calculation. Real NAV will differ slightly.
- **The constituent list is static** and needs manual updates when the index changes.
- **Yahoo Finance data is unofficial.** `yfinance` is not affiliated with NSE, and its data can be delayed, incomplete or revised.
- **Equal weighting** means position sizes ignore volatility and correlation.
- **Single-file design.** There are no automated tests in this repository yet.

---

## Roadmap Ideas

- Add unit tests for the ranking, buffer and regime logic.
- Read the constituent list from NSE's official index file rather than hard-coding it.
- Add a `--dry-run` flag that writes to a separate workbook.
- Model transaction costs and minimum lot sizes.
- Export the action table as CSV for broker import.

---

## License

Choose a licence before publishing (for example MIT or Apache-2.0) and add a `LICENSE` file. This README does not grant any licence by itself.
