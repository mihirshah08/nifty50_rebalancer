#!/usr/bin/env python3
"""
Nifty 50 Monthly Portfolio Rebalancer (Excel-Driven, Standalone Script)
=======================================================================
Implements all 4 backtested strategy enhancements:
  1. Top-10 Relative Strength Concentration: Ranks Nifty 50 stocks trading above
     their 200 DMA by % distance above 200 DMA ((LTP - 200 DMA) / 200 DMA) and
     concentrates capital into the Top 10 leaders.
  2. Strict < 200 DMA Exit Rule: Immediately sells any holding whose LTP falls
     below or equal to its 200 DMA (LTP <= 200 DMA) or exits the Nifty 50 index.
  3. Top-15 Rank-Retention Buffer: Retains existing holdings from last month's
     Excel sheet as long as they remain above their 200 DMA and rank within the
     Top 15 (preventing unnecessary monthly churn and short-term tax drag).
  4. Market-Breadth Cash / Liquid-ETF Filter: When >= 50% (>= 25 of 50) of
     Nifty 50 stocks trade above their 200 DMA (Bull Regime), allocates 100% to
     equities (10% per stock). When < 50% (< 25 stocks) trade above their 200 DMA
     (Defensive Regime), scales equity exposure to 50% (5% per stock) and parks
     50% in Nippon India ETF Liquid BeES (LIQUIDBEES).

Monthly Workflow (Zero Manual Entry):
-------------------------------------
  • Run `python nifty50_monthly_rebalancer.py` on the 1st of every month.
  • The script automatically reads your previous month's holdings, share counts,
    and leftover cash from `Nifty50_Portfolio_Tracker.xlsx` (Sheet: `Current_Holdings`).
  • It fetches live NSE prices & 200 DMAs via `yfinance` (when `--live` is passed)
    OR reads the updated LTP & 200 DMA columns directly from the Excel file
    (`Nifty50_Screener` sheet or your existing LTP/200DMA sheet).
  • It prints the clean ADD / RETAIN / REMOVE action table in the terminal and
    saves the updated `Nifty50_Portfolio_Tracker.xlsx` ready for next month!
"""

import argparse
from datetime import datetime
import math
import os
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# Official Nifty 50 Constituents & Baseline November 2026 LTP / 200 DMA (Google Finance / NSE)
DEFAULT_NIFTY50_DATA = [
    ("ADANIENT", "Adani Enterprises Ltd", 2685.00, 2624.57),
    ("ADANIPORTS", "Adani Ports and SEZ Ltd", 1919.70, 1675.43),
    ("APOLLOHOSP", "Apollo Hospitals Enterprise Ltd", 8396.50, 8205.90),
    ("ASIANPAINT", "Asian Paints Ltd", 2493.50, 2404.90),
    ("AXISBANK", "Axis Bank Ltd", 1282.00, 1219.53),
    ("BAJAJ-AUTO", "Bajaj Auto Ltd", 10820.00, 10309.26),
    ("BAJFINANCE", "Bajaj Finance Ltd", 1052.10, 991.50),
    ("BAJAJFINSV", "Bajaj Finserv Ltd", 1823.50, 1730.89),
    ("BEL", "Bharat Electronics Ltd", 387.40, 396.79),
    ("BHARTIARTL", "Bharti Airtel Ltd", 2082.00, 1938.80),
    ("BSE", "BSE Ltd", 3761.00, 3380.89),
    ("CIPLA", "Cipla Ltd", 1402.30, 1347.82),
    ("COALINDIA", "Coal India Ltd", 396.20, 381.47),
    ("DRREDDY", "Dr. Reddy's Laboratories Ltd", 1368.60, 1295.01),
    ("EICHERMOT", "Eicher Motors Ltd", 7557.50, 7126.90),
    ("ETERNAL", "Eternal Ltd (Zomato)", 354.75, 280.27),
    ("GRASIM", "Grasim Industries Ltd", 3186.50, 2975.80),
    ("HCLTECH", "HCL Technologies Ltd", 1439.50, 1296.69),
    ("HDFCBANK", "HDFC Bank Ltd", 773.70, 811.83),
    ("HDFCLIFE", "HDFC Life Insurance Co Ltd", 550.00, 616.79),
    ("HINDALCO", "Hindalco Industries Ltd", 920.50, 864.11),
    ("HINDUNILVR", "Hindustan Unilever Ltd", 1976.80, 2180.66),
    ("ICICIBANK", "ICICI Bank Ltd", 1527.60, 1352.09),
    ("INDIGO", "InterGlobe Aviation Ltd (IndiGo)", 5380.40, 4756.44),
    ("INFY", "Infosys Ltd", 1168.50, 1345.32),
    ("ITC", "ITC Ltd", 280.00, 336.56),
    ("JIOFIN", "Jio Financial Services Ltd", 226.20, 254.08),
    ("JSWSTEEL", "JSW Steel Ltd", 1341.60, 1244.41),
    ("KOTAKBANK", "Kotak Mahindra Bank Ltd", 472.55, 399.85),
    ("LT", "Larsen & Toubro Ltd", 4051.40, 3819.99),
    ("M&M", "Mahindra & Mahindra Ltd", 3233.20, 3356.11),
    ("MARUTI", "Maruti Suzuki India Ltd", 12625.00, 13180.42),
    ("MAXHEALTH", "Max Healthcare Institute Ltd", 924.25, 976.43),
    ("NESTLEIND", "Nestle India Ltd", 1258.30, 1325.84),
    ("NTPC", "NTPC Ltd", 317.80, 344.15),
    ("ONGC", "Oil and Natural Gas Corp Ltd", 230.25, 255.48),
    ("POWERGRID", "Power Grid Corp of India Ltd", 254.85, 248.12),
    ("RELIANCE", "Reliance Industries Ltd", 1310.70, 1358.92),
    ("SBILIFE", "SBI Life Insurance Co Ltd", 1718.70, 1658.40),
    ("SBIN", "State Bank of India", 998.50, 912.60),
    ("SHRIRAMFIN", "Shriram Finance Ltd", 932.45, 968.20),
    ("SUNPHARMA", "Sun Pharmaceutical Ind Ltd", 1847.30, 1808.68),
    ("TATACONSUM", "Tata Consumer Products Ltd", 1081.20, 1126.40),
    ("TMPV", "Tata Motors Passenger Vehicles Ltd", 286.85, 342.10),
    ("TATASTEEL", "Tata Steel Ltd", 180.66, 188.45),
    ("TCS", "Tata Consultancy Services Ltd", 2504.30, 2418.90),
    ("TECHM", "Tech Mahindra Ltd", 1826.00, 1506.39),
    ("TITAN", "Titan Company Ltd", 4950.20, 4470.83),
    ("TRENT", "Trent Ltd", 3042.70, 2912.50),
    ("ULTRACEMCO", "UltraTech Cement Ltd", 12254.00, 11680.25),
]

# Previous Month's (October 2026) Baseline Holdings (used only if no Excel file exists yet)
INITIAL_OCT2026_HOLDINGS = [
    ("ETERNAL", 460, 354.75),
    ("TECHM", 89, 1826.00),
    ("KOTAKBANK", 345, 472.55),
    ("ADANIPORTS", 85, 1919.70),
    ("TITAN", 33, 4950.20),
    ("GRASIM", 51, 3186.50),
    ("BAJAJ-AUTO", 15, 10820.00),
    ("APOLLOHOSP", 19, 8396.50),
    ("ADANIENT", 60, 2685.00),
    ("SUNPHARMA", 88, 1847.30),
]
INITIAL_OCT2026_CASH = 8130.62


def auto_fit_sheet_columns(ws, skip_rows=(1,), min_w=12, max_w=36):
    """Dynamically adjusts column widths ignoring header/merged rows."""
    merged_multi_col = {
        cell.coordinate
        for rng in ws.merged_cells.ranges
        if rng.min_col != rng.max_col
        for row in ws[rng.coord]
        for cell in row
    }
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in skip_rows or cell.coordinate in merged_multi_col:
                continue
            val = str(cell.value) if cell.value is not None else ""
            if val.startswith("="):
                max_len = max(max_len, 14)
            else:
                max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, min_w), max_w)


def load_previous_month_excel(excel_path):
    """
    Reads the previous month's Excel workbook (`Nifty50_Portfolio_Tracker.xlsx` or custom sheet)
    to extract:
      1. `prev_holdings`: dict of {Symbol: Held_Shares}
      2. `prev_liquid_shares`: int shares of LIQUIDBEES held
      3. `prev_cash`: float uninvested cash balance (INR)
      4. `screener_df`: DataFrame of Nifty 50 stocks with Symbol, Company_Name, LTP, DMA_200
    """
    if not os.path.exists(excel_path):
        return None, 0, 0.0, None

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    wb_f = openpyxl.load_workbook(excel_path, data_only=False)

    prev_holdings = {}
    prev_liquid_shares = 0
    prev_cash = 0.0

    # 1. Read Current_Holdings tab if present
    if "Current_Holdings" in wb.sheetnames:
        ws_h = wb["Current_Holdings"]
        ws_hf = wb_f["Current_Holdings"]
        for r in range(2, ws_h.max_row + 1):
            sym = ws_h.cell(r, 1).value or ws_hf.cell(r, 1).value
            sh = ws_h.cell(r, 2).value
            if sh is None:
                sh = ws_hf.cell(r, 2).value
            ltp_val = ws_h.cell(r, 3).value
            if ltp_val is None:
                ltp_val = ws_hf.cell(r, 3).value
            if not sym:
                continue
            sym = str(sym).strip().upper()
            if sym in ("SYMBOL", "TOTAL", "PORTFOLIO TOTAL", ""):
                continue
            if sym == "CASH_BALANCE":
                try:
                    prev_cash = float(ltp_val or 0.0)
                except (ValueError, TypeError):
                    prev_cash = 0.0
            elif sym.startswith("LIQUIDBEES"):
                try:
                    prev_liquid_shares = int(float(sh or 0))
                except (ValueError, TypeError):
                    prev_liquid_shares = 0
            else:
                try:
                    sh_int = int(float(sh or 0))
                    if sh_int > 0:
                        prev_holdings[sym] = sh_int
                except (ValueError, TypeError):
                    pass

    # 2. Read Nifty50_Screener (or fallback first sheet with LTP & 200 DMA)
    screener_rows = []
    target_sheet = "Nifty50_Screener" if "Nifty50_Screener" in wb.sheetnames else wb.sheetnames[0]
    ws_s = wb[target_sheet]
    ws_sf = wb_f[target_sheet]

    header_row = None
    col_map = {}
    for r in range(1, min(15, ws_s.max_row + 1)):
        row_vals = [str(ws_s.cell(r, c).value or ws_sf.cell(r, c).value or "").strip().upper() for c in range(1, ws_s.max_column + 1)]
        if any(k in row_vals for k in ("SYMBOL", "STOCK", "TICKER")) and any("LTP" in v for v in row_vals):
            header_row = r
            for c_idx, v in enumerate(row_vals, start=1):
                if v in ("SYMBOL", "STOCK", "TICKER"):
                    col_map["sym"] = c_idx
                elif "COMPANY" in v or "NAME" in v:
                    col_map["name"] = c_idx
                elif "LTP" in v or "PRICE" in v:
                    col_map["ltp"] = c_idx
                elif "200" in v or "DMA" in v:
                    col_map["dma"] = c_idx
                elif "PORTFOLIO" in v or "IN_PORTFOLIO" in v:
                    col_map["port"] = c_idx
            break

    name_lookup = {s: n for s, n, _, _ in DEFAULT_NIFTY50_DATA}
    if header_row is not None and "sym" in col_map and "ltp" in col_map and "dma" in col_map:
        for r in range(header_row + 1, ws_s.max_row + 1):
            sym_val = ws_s.cell(r, col_map["sym"]).value or ws_sf.cell(r, col_map["sym"]).value
            if not sym_val:
                continue
            sym_str = str(sym_val).strip().upper().replace("NSE:", "")
            if sym_str in ("TOTAL", "AVERAGE", "AVERAGE / COUNT", "SUMMARY", ""):
                continue
            ltp_v = ws_s.cell(r, col_map["ltp"]).value
            if ltp_v is None:
                ltp_v = ws_sf.cell(r, col_map["ltp"]).value
            dma_v = ws_s.cell(r, col_map["dma"]).value
            if dma_v is None:
                dma_v = ws_sf.cell(r, col_map["dma"]).value
            try:
                ltp_f = float(ltp_v)
                dma_f = float(dma_v)
            except (ValueError, TypeError):
                continue
            c_name = name_lookup.get(sym_str, sym_str)
            if "name" in col_map:
                n_val = ws_s.cell(r, col_map["name"]).value or ws_sf.cell(r, col_map["name"]).value
                if n_val:
                    c_name = str(n_val).strip()
            screener_rows.append((sym_str, c_name, ltp_f, dma_f))

            if not prev_holdings and "port" in col_map:
                p_flag = str(ws_s.cell(r, col_map["port"]).value or "").strip().upper()
                if p_flag == "YES":
                    prev_holdings[sym_str] = 1

    wb.close()
    wb_f.close()
    screener_df = pd.DataFrame(screener_rows, columns=["Symbol", "Company_Name", "LTP", "DMA_200"]) if screener_rows else None
    return prev_holdings, prev_liquid_shares, prev_cash, screener_df


def try_fetch_live_yfinance(screener_df):
    """
    If `--live` is passed on a machine with internet & yfinance installed,
    updates LTP and 200 DMA for all 50 Nifty 50 stocks from NSE (`SYMBOL.NS`).
    """
    try:
        import yfinance as yf
        tickers = screener_df["Symbol"].tolist()
        ns_map = {t: f"{t}.NS" for t in tickers}
        data = yf.download(list(ns_map.values()), period="1y", interval="1d", progress=False)["Close"]
        if data is None or data.empty:
            return screener_df, False
        updated_rows = []
        for _, row in screener_df.iterrows():
            t = row["Symbol"]
            yf_t = ns_map[t]
            ltp_val = float(row["LTP"])
            dma_val = float(row["DMA_200"])
            if yf_t in data.columns:
                s = data[yf_t].dropna()
                if len(s) >= 200:
                    ltp_val = float(s.iloc[-1])
                    dma_val = float(s.iloc[-200:].mean())
            updated_rows.append((t, row["Company_Name"], round(ltp_val, 2), round(dma_val, 2)))
        return pd.DataFrame(updated_rows, columns=["Symbol", "Company_Name", "LTP", "DMA_200"]), True
    except Exception:
        return screener_df, False


def run_rebalance_calculation(
    screener_df,
    prev_holdings,
    prev_liquid_shares=0,
    prev_cash=0.0,
    override_capital=None,
    top_n=10,
    keep_rank=15,
    breadth_threshold=0.50,
    defensive_equity_wt=0.50,
):
    """
    Computes the current portfolio value from previous month's Excel holdings
    and generates the new month's Add / Retain / Remove & Weightage table.
    """
    df = screener_df.copy()
    df["Dist_200DMA_Pct"] = ((df["LTP"] / df["DMA_200"]) - 1.0) * 100.0
    df["Above_200DMA"] = df["LTP"] > df["DMA_200"]

    above_df = df[df["Above_200DMA"]].sort_values(by="Dist_200DMA_Pct", ascending=False).reset_index(drop=True)
    above_df["RS_Rank"] = above_df.index + 1

    rank_map = dict(zip(above_df["Symbol"], above_df["RS_Rank"]))
    ltp_map = dict(zip(df["Symbol"], df["LTP"]))
    dma_map = dict(zip(df["Symbol"], df["DMA_200"]))
    name_map = dict(zip(df["Symbol"], df["Company_Name"]))
    dist_map = dict(zip(df["Symbol"], df["Dist_200DMA_Pct"]))
    active_universe = set(df["Symbol"].tolist())

    if override_capital is not None:
        total_capital = float(override_capital)
    elif prev_holdings:
        eq_val = sum(sh * ltp_map.get(sym, 0.0) for sym, sh in prev_holdings.items())
        liq_val = prev_liquid_shares * 1000.00
        total_capital = eq_val + liq_val + prev_cash
    else:
        total_capital = 1000000.00

    num_above = len(above_df)
    total_stocks = len(df)
    breadth_ratio = num_above / total_stocks if total_stocks > 0 else 0.0
    is_bull_regime = breadth_ratio >= breadth_threshold
    target_eq_wt = 1.00 if is_bull_regime else defensive_equity_wt
    target_liq_wt = 1.00 - target_eq_wt

    retained = []
    removed_rows = []
    for sym, prev_sh in prev_holdings.items():
        p = ltp_map.get(sym, 0.0)
        d = dma_map.get(sym, 0.0)
        rk = rank_map.get(sym, 50)
        if sym not in active_universe:
            removed_rows.append((sym, "REJECTED", prev_sh, p, d, rk, "Excluded from Nifty 50 Index"))
        elif p <= d:
            removed_rows.append((sym, "REJECTED", prev_sh, p, d, 50, "Fell Below 200 DMA (Strict Exit)"))
        elif rk > keep_rank:
            removed_rows.append((sym, "REJECTED", prev_sh, p, d, rk, f"Rank #{rk} Outside Top-{keep_rank} Buffer"))
        else:
            retained.append(sym)

    retained = sorted(retained, key=lambda s: rank_map[s])
    if len(retained) > top_n:
        for sym in retained[top_n:]:
            removed_rows.append((sym, "REJECTED", prev_holdings[sym], ltp_map[sym], dma_map[sym], rank_map[sym], f"Trimmed to Top-{top_n} Cap"))
        retained = retained[:top_n]

    slots_open = max(0, top_n - len(retained))
    candidates = [s for s in above_df["Symbol"].tolist() if s not in retained]
    added = candidates[:slots_open]

    final_equities = retained + added
    n_eq = len(final_equities)
    per_stock_wt = (target_eq_wt / n_eq) if n_eq > 0 else 0.0
    if n_eq == 0:
        target_liq_wt = 1.00
        target_eq_wt = 0.00

    rebalance_rows = []
    for sym in final_equities:
        is_new = sym in added
        act = "ADD (BUY)" if is_new else "RETAIN (HOLD)"
        rk = int(rank_map[sym])
        p = float(ltp_map[sym])
        d = float(dma_map[sym])
        prev_sh = int(prev_holdings.get(sym, 0))
        t_val = total_capital * per_stock_wt
        t_sh = int(math.floor(t_val / p)) if p > 0 else 0
        reason = f"Top-{top_n} RS Leader (Rank #{rk})" if is_new else f"Retained in Top-{keep_rank} Buffer (Rank #{rk})"
        rebalance_rows.append({
            "Symbol": sym,
            "Company_Name": name_map.get(sym, sym),
            "Action": act,
            "RS_Rank": rk,
            "LTP": p,
            "DMA_200": d,
            "Dist_Pct": dist_map[sym],
            "Prev_Shares": prev_sh,
            "Target_Weight": per_stock_wt,
            "Target_Value": t_val,
            "Target_Shares": t_sh,
            "Trade_Shares": t_sh - prev_sh,
            "Actual_Invested": t_sh * p,
            "Reason": reason,
            "Source_URL": f"https://www.google.com/finance/quote/{sym}:NSE",
        })

    rebalance_rows = sorted(rebalance_rows, key=lambda x: x["RS_Rank"])

    for sym, _, prev_sh, p, d, rk, reason in removed_rows:
        rebalance_rows.append({
            "Symbol": sym,
            "Company_Name": name_map.get(sym, sym),
            "Action": "REMOVE (SELL)",
            "RS_Rank": int(rk),
            "LTP": p,
            "DMA_200": d,
            "Dist_Pct": ((p / d) - 1.0) * 100.0 if d > 0 else 0.0,
            "Prev_Shares": int(prev_sh),
            "Target_Weight": 0.0,
            "Target_Value": 0.0,
            "Target_Shares": 0,
            "Trade_Shares": -int(prev_sh),
            "Actual_Invested": 0.0,
            "Reason": reason,
            "Source_URL": f"https://www.google.com/finance/quote/{sym}:NSE",
        })

    liq_target_val = total_capital * target_liq_wt
    liq_target_sh = int(math.floor(liq_target_val / 1000.0))
    rebalance_rows.append({
        "Symbol": "LIQUIDBEES",
        "Company_Name": "Nippon India ETF Liquid BeES (Cash/Liquid Proxy)",
        "Action": "ADD (BUY)" if liq_target_sh > prev_liquid_shares else ("RETAIN (HOLD)" if liq_target_sh > 0 else "HOLD (0% CASH)"),
        "RS_Rank": 0,
        "LTP": 1000.00,
        "DMA_200": 1000.00,
        "Dist_Pct": 0.0,
        "Prev_Shares": int(prev_liquid_shares),
        "Target_Weight": target_liq_wt,
        "Target_Value": liq_target_val,
        "Target_Shares": liq_target_sh,
        "Trade_Shares": liq_target_sh - int(prev_liquid_shares),
        "Actual_Invested": liq_target_sh * 1000.0,
        "Reason": f"Breadth {breadth_ratio*100:.1f}% (<{breadth_threshold*100:.0f}% Defensive)" if target_liq_wt > 0 else f"Breadth {breadth_ratio*100:.1f}% (>={breadth_threshold*100:.0f}% Bull Regime)",
        "Source_URL": "https://www.google.com/finance/quote/LIQUIDBEES:NSE",
    })

    rebalance_df = pd.DataFrame(rebalance_rows)
    total_invested_after = float(rebalance_df["Actual_Invested"].sum())
    new_cash_balance = round(total_capital - total_invested_after, 2)

    df["RS_Rank"] = df["Symbol"].map(lambda s: int(rank_map[s]) if s in rank_map else 50)
    df["In_Portfolio"] = df["Symbol"].map(lambda s: "YES" if s in final_equities else "NO")
    df = df.sort_values(by=["Above_200DMA", "Dist_200DMA_Pct"], ascending=[False, False]).reset_index(drop=True)

    summary = {
        "as_of_date": datetime.now().strftime("%Y-%m-%d"),
        "total_capital": total_capital,
        "total_invested": total_invested_after,
        "new_cash_balance": new_cash_balance,
        "num_above_200dma": num_above,
        "total_universe": total_stocks,
        "breadth_pct": breadth_ratio * 100.0,
        "regime": "BULL REGIME (100% Equity)" if is_bull_regime else f"DEFENSIVE REGIME ({target_eq_wt*100:.0f}% Equity / {target_liq_wt*100:.0f}% Liquid ETF)",
        "num_retained": len(retained),
        "num_added": len(added),
        "num_removed": len(removed_rows),
        "per_stock_wt_pct": per_stock_wt * 100.0,
        "liquid_wt_pct": target_liq_wt * 100.0,
        "final_equities": final_equities,
        "liquid_shares": liq_target_sh,
    }
    return rebalance_df, df, summary


def write_rebalance_excel(excel_path, rebalance_df, screener_df, summary):
    """
    Writes the multi-sheet Excel workbook (`Nifty50_Portfolio_Tracker.xlsx`):
      1. `Rebalance_Action_Sheet` — Actionable ADD / RETAIN / REMOVE table + Embedded BarChart
      2. `Current_Holdings`       — Active holdings & share counts automatically read by next month's run
      3. `Nifty50_Screener`       — All 50 Nifty 50 stocks with LTP, 200 DMA, % vs 200 DMA, and Portfolio status
    """
    wb = openpyxl.Workbook()

    NAVY_HEADER = "1B365D"
    ALT_ROW = "F4F7FA"
    BORDER_HEX = "D9E2EC"
    ADD_FILL = "D1E7DD"
    RETAIN_FILL = "CFF4FC"
    REMOVE_FILL = "F8D7DA"

    font_tbl_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=10, color="1A202C")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1A202C")
    font_link = Font(name="Calibri", size=10, color="1D4ED8", underline="single")

    fill_hdr = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    fill_alt = PatternFill(start_color=ALT_ROW, end_color=ALT_ROW, fill_type="solid")
    fill_add = PatternFill(start_color=ADD_FILL, end_color=ADD_FILL, fill_type="solid")
    fill_retain = PatternFill(start_color=RETAIN_FILL, end_color=RETAIN_FILL, fill_type="solid")
    fill_remove = PatternFill(start_color=REMOVE_FILL, end_color=REMOVE_FILL, fill_type="solid")

    thin_side = Side(style="thin", color=BORDER_HEX)
    double_navy = Side(style="double", color=NAVY_HEADER)

    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    total_border = Border(top=thin_side, bottom=double_navy, left=thin_side, right=thin_side)

    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")

    # =========================================================================
    # SHEET 1: Rebalance_Action_Sheet (Row 1 Header for 100% Column Homogeneity)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "Rebalance_Action_Sheet"
    ws1.views.sheetView[0].showGridLines = True
    ws1.page_setup.orientation = "landscape"
    ws1.sheet_properties.pageSetUpPr.fitToPage = True

    headers_s1 = [
        "Symbol",
        "Company Name",
        "Rebalance Action",
        "RS Rank",
        "LTP (INR)",
        "200 DMA (INR)",
        "% vs 200 DMA",
        "Prev Shares",
        "Target Weight",
        "Target Value (INR)",
        "Target Shares",
        "Trade Shares (+/-)",
        "Selection Reason",
        "Source URL",
    ]
    ws1.row_dimensions[1].height = 25
    for col_idx, h_text in enumerate(headers_s1, start=1):
        c = ws1.cell(row=1, column=col_idx, value=h_text)
        c.font = font_tbl_hdr
        c.fill = fill_hdr
        c.border = cell_border
        c.alignment = align_left if col_idx in (1, 2, 3, 13, 14) else align_right

    start_row_s1 = 2
    cap_val = round(float(summary["total_capital"]), 2)
    for idx, r in rebalance_df.iterrows():
        row_num = start_row_s1 + idx
        ws1.row_dimensions[row_num].height = 20
        act = str(r["Action"])
        url = str(r["Source_URL"])

        ws1.cell(row=row_num, column=1, value=str(r["Symbol"]))
        ws1.cell(row=row_num, column=2, value=str(r["Company_Name"]))
        ws1.cell(row=row_num, column=3, value=act)
        ws1.cell(row=row_num, column=4, value=int(r["RS_Rank"]))
        ws1.cell(row=row_num, column=5, value=round(float(r["LTP"]), 2))
        ws1.cell(row=row_num, column=6, value=round(float(r["DMA_200"]), 2))
        ws1.cell(row=row_num, column=7, value=f"=(E{row_num}-F{row_num})/F{row_num}")
        ws1.cell(row=row_num, column=8, value=int(r["Prev_Shares"]))
        ws1.cell(row=row_num, column=9, value=round(float(r["Target_Weight"]), 4))
        ws1.cell(row=row_num, column=10, value=f"={cap_val}*I{row_num}")
        ws1.cell(row=row_num, column=11, value=f"=INT(J{row_num}/E{row_num})")
        ws1.cell(row=row_num, column=12, value=f"=K{row_num}-H{row_num}")
        ws1.cell(row=row_num, column=13, value=str(r["Reason"]))
        c_link = ws1.cell(row=row_num, column=14, value=f"NSE: {r['Symbol']}")
        c_link.hyperlink = url

        row_fill = fill_alt if row_num % 2 == 0 else None
        for col_idx in range(1, 15):
            c = ws1.cell(row=row_num, column=col_idx)
            c.font = font_bold if col_idx == 1 else (font_link if col_idx == 14 else font_data)
            c.border = cell_border
            if col_idx == 3:
                c.alignment = align_left
                if "ADD" in act:
                    c.fill = fill_add
                elif "RETAIN" in act:
                    c.fill = fill_retain
                elif "REMOVE" in act:
                    c.fill = fill_remove
                elif row_fill:
                    c.fill = row_fill
            else:
                if row_fill:
                    c.fill = row_fill
                c.alignment = align_left if col_idx in (1, 2, 3, 13, 14) else align_right

        ws1.cell(row=row_num, column=4).number_format = "#,##0"
        ws1.cell(row=row_num, column=5).number_format = "₹#,##0.00"
        ws1.cell(row=row_num, column=6).number_format = "₹#,##0.00"
        ws1.cell(row=row_num, column=7).number_format = "+0.00%;-0.00%;0.00%"
        ws1.cell(row=row_num, column=8).number_format = "#,##0"
        ws1.cell(row=row_num, column=9).number_format = "0.00%"
        ws1.cell(row=row_num, column=10).number_format = "₹#,##0.00"
        ws1.cell(row=row_num, column=11).number_format = "#,##0"
        ws1.cell(row=row_num, column=12).number_format = "+#,##0;-#,##0;0"

    end_row_s1 = start_row_s1 + len(rebalance_df) - 1
    tot_row_s1 = end_row_s1 + 1
    ws1.row_dimensions[tot_row_s1].height = 22

    ws1.cell(row=tot_row_s1, column=1, value="Total")
    ws1.cell(row=tot_row_s1, column=2, value=f"Regime: {summary['regime']} | Breadth: {summary['breadth_pct']:.1f}%")
    ws1.cell(row=tot_row_s1, column=3, value=f"{summary['num_retained']} Hold, {summary['num_added']} Buy, {summary['num_removed']} Sell")
    ws1.cell(row=tot_row_s1, column=4, value=f"=COUNT(D{start_row_s1}:D{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=5, value=f"=AVERAGE(E{start_row_s1}:E{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=6, value=f"=AVERAGE(F{start_row_s1}:F{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=7, value=f"=AVERAGE(G{start_row_s1}:G{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=8, value=f"=SUM(H{start_row_s1}:H{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=9, value=f"=SUM(I{start_row_s1}:I{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=10, value=f"=SUM(J{start_row_s1}:J{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=11, value=f"=SUM(K{start_row_s1}:K{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=12, value=f"=SUM(L{start_row_s1}:L{end_row_s1})")
    ws1.cell(row=tot_row_s1, column=13, value=f"Residual Cash: ₹{summary['new_cash_balance']:,.2f}")
    c_tot_url = ws1.cell(row=tot_row_s1, column=14, value="NSE Nifty 50 Tracker")
    c_tot_url.hyperlink = "https://www.nseindia.com/index-tracker/NIFTY%2050"

    for col_idx in range(1, 15):
        c = ws1.cell(row=tot_row_s1, column=col_idx)
        c.font = font_bold if col_idx != 14 else font_link
        c.border = total_border
        c.alignment = align_left if col_idx in (1, 2, 3, 13, 14) else align_right

    ws1.cell(row=tot_row_s1, column=4).number_format = "#,##0"
    ws1.cell(row=tot_row_s1, column=5).number_format = "₹#,##0.00"
    ws1.cell(row=tot_row_s1, column=6).number_format = "₹#,##0.00"
    ws1.cell(row=tot_row_s1, column=7).number_format = "+0.00%;-0.00%;0.00%"
    ws1.cell(row=tot_row_s1, column=8).number_format = "#,##0"
    ws1.cell(row=tot_row_s1, column=9).number_format = "0.00%"
    ws1.cell(row=tot_row_s1, column=10).number_format = "₹#,##0.00"
    ws1.cell(row=tot_row_s1, column=11).number_format = "#,##0"
    ws1.cell(row=tot_row_s1, column=12).number_format = "+#,##0;-#,##0;0"

    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Relative Strength (% Distance vs 200 DMA) Across Rebalance Candidates"
    chart1.y_axis.title = "% Distance vs 200 DMA"
    chart1.x_axis.title = "Symbol"
    chart1.width = 22
    chart1.height = 10
    data_ref1 = Reference(ws1, min_col=7, min_row=1, max_col=7, max_row=end_row_s1)
    cats_ref1 = Reference(ws1, min_col=1, min_row=start_row_s1, max_row=end_row_s1)
    chart1.add_data(data_ref1, titles_from_data=True)
    chart1.set_categories(cats_ref1)
    chart1.legend = None
    ws1.add_chart(chart1, f"A{tot_row_s1 + 3}")

    ws1.freeze_panes = "A2"
    auto_fit_sheet_columns(ws1, skip_rows=(1,))

    # =========================================================================
    # SHEET 2: Current_Holdings (Automatically read by next month's script run!)
    # =========================================================================
    ws2 = wb.create_sheet(title="Current_Holdings")
    ws2.views.sheetView[0].showGridLines = True
    ws2.page_setup.orientation = "landscape"
    ws2.sheet_properties.pageSetUpPr.fitToPage = True

    headers_s2 = [
        "Symbol",
        "Held_Shares",
        "Entry_Rebalance_LTP",
        "Allocated_Value_INR",
        "Portfolio_Weight",
        "Status",
        "Source_URL",
    ]
    ws2.row_dimensions[1].height = 25
    for col_idx, h_text in enumerate(headers_s2, start=1):
        c = ws2.cell(row=1, column=col_idx, value=h_text)
        c.font = font_tbl_hdr
        c.fill = fill_hdr
        c.border = cell_border
        c.alignment = align_left if col_idx in (1, 7) else (align_center if col_idx == 6 else align_right)

    active_df = rebalance_df[rebalance_df["Target_Shares"] > 0].reset_index(drop=True)
    start_r2 = 2
    curr_r2 = start_r2
    for _, r in active_df.iterrows():
        ws2.row_dimensions[curr_r2].height = 20
        sym = str(r["Symbol"])
        sh = int(r["Target_Shares"])
        ltp_v = round(float(r["LTP"]), 2)
        wt_v = round(float(r["Target_Weight"]), 4)
        url = str(r["Source_URL"])

        ws2.cell(row=curr_r2, column=1, value=sym)
        ws2.cell(row=curr_r2, column=2, value=sh)
        ws2.cell(row=curr_r2, column=3, value=ltp_v)
        ws2.cell(row=curr_r2, column=4, value=f"=B{curr_r2}*C{curr_r2}")
        ws2.cell(row=curr_r2, column=5, value=wt_v)
        ws2.cell(row=curr_r2, column=6, value="ACTIVE")
        c_u = ws2.cell(row=curr_r2, column=7, value=f"NSE: {sym}")
        c_u.hyperlink = url
        curr_r2 += 1

    ws2.row_dimensions[curr_r2].height = 20
    ws2.cell(row=curr_r2, column=1, value="CASH_BALANCE")
    ws2.cell(row=curr_r2, column=2, value=1)
    ws2.cell(row=curr_r2, column=3, value=round(float(summary["new_cash_balance"]), 2))
    ws2.cell(row=curr_r2, column=4, value=f"=B{curr_r2}*C{curr_r2}")
    ws2.cell(row=curr_r2, column=5, value=round(float(summary["new_cash_balance"]) / summary["total_capital"], 4) if summary["total_capital"] > 0 else 0.0)
    ws2.cell(row=curr_r2, column=6, value="ACTIVE")
    c_cash_u = ws2.cell(row=curr_r2, column=7, value="NSE Nifty 50 Tracker")
    c_cash_u.hyperlink = "https://www.nseindia.com/index-tracker/NIFTY%2050"

    end_r2 = curr_r2
    for r_idx in range(start_r2, end_r2 + 1):
        row_fill = fill_alt if r_idx % 2 == 0 else None
        for c_idx in range(1, 8):
            c = ws2.cell(row=r_idx, column=c_idx)
            c.font = font_bold if c_idx == 1 else (font_link if c_idx == 7 else font_data)
            c.border = cell_border
            if row_fill:
                c.fill = row_fill
            c.alignment = align_left if c_idx in (1, 7) else (align_center if c_idx == 6 else align_right)
        ws2.cell(row=r_idx, column=2).number_format = "#,##0"
        ws2.cell(row=r_idx, column=3).number_format = "₹#,##0.00"
        ws2.cell(row=r_idx, column=4).number_format = "₹#,##0.00"
        ws2.cell(row=r_idx, column=5).number_format = "0.00%"

    tot_r2 = end_r2 + 1
    ws2.row_dimensions[tot_r2].height = 22
    ws2.cell(row=tot_r2, column=1, value="Total")
    ws2.cell(row=tot_r2, column=2, value=f"=SUM(B{start_r2}:B{end_r2})")
    ws2.cell(row=tot_r2, column=3, value=f"=AVERAGE(C{start_r2}:C{end_r2})")
    ws2.cell(row=tot_r2, column=4, value=f"=SUM(D{start_r2}:D{end_r2})")
    ws2.cell(row=tot_r2, column=5, value=f"=SUM(E{start_r2}:E{end_r2})")
    ws2.cell(row=tot_r2, column=6, value="ACTIVE")
    c_tot2_u = ws2.cell(row=tot_r2, column=7, value="NSE Nifty 50 Tracker")
    c_tot2_u.hyperlink = "https://www.nseindia.com/index-tracker/NIFTY%2050"

    for c_idx in range(1, 8):
        c = ws2.cell(row=tot_r2, column=c_idx)
        c.font = font_bold if c_idx != 7 else font_link
        c.border = total_border
        c.alignment = align_left if c_idx in (1, 7) else (align_center if c_idx == 6 else align_right)
    ws2.cell(row=tot_r2, column=2).number_format = "#,##0"
    ws2.cell(row=tot_r2, column=3).number_format = "₹#,##0.00"
    ws2.cell(row=tot_r2, column=4).number_format = "₹#,##0.00"
    ws2.cell(row=tot_r2, column=5).number_format = "0.00%"

    chart2 = BarChart()
    chart2.type = "col"
    chart2.style = 10
    chart2.title = "Allocated Capital (INR) by Active Holding"
    chart2.y_axis.title = "Allocated Value (INR)"
    chart2.x_axis.title = "Holding"
    chart2.width = 18
    chart2.height = 9
    data_ref2 = Reference(ws2, min_col=4, min_row=1, max_col=4, max_row=end_r2)
    cats_ref2 = Reference(ws2, min_col=1, min_row=start_r2, max_row=end_r2)
    chart2.add_data(data_ref2, titles_from_data=True)
    chart2.set_categories(cats_ref2)
    chart2.legend = None
    ws2.add_chart(chart2, f"A{tot_r2 + 3}")

    ws2.freeze_panes = "A2"
    auto_fit_sheet_columns(ws2, skip_rows=(1,))

    # =========================================================================
    # SHEET 3: Nifty50_Screener (All 50 Nifty 50 Constituents LTP & 200 DMA)
    # =========================================================================
    ws3 = wb.create_sheet(title="Nifty50_Screener")
    ws3.views.sheetView[0].showGridLines = True
    ws3.page_setup.orientation = "landscape"
    ws3.sheet_properties.pageSetUpPr.fitToPage = True

    headers_s3 = [
        "Symbol",
        "Company_Name",
        "LTP",
        "DMA_200",
        "Dist_200DMA_Pct",
        "Above_200DMA",
        "RS_Rank",
        "In_Portfolio",
        "Source_URL",
    ]
    ws3.row_dimensions[1].height = 25
    for col_idx, h_text in enumerate(headers_s3, start=1):
        c = ws3.cell(row=1, column=col_idx, value=h_text)
        c.font = font_tbl_hdr
        c.fill = fill_hdr
        c.border = cell_border
        c.alignment = align_left if col_idx in (1, 2, 9) else (align_center if col_idx in (6, 8) else align_right)

    start_r3 = 2
    for idx, r in screener_df.iterrows():
        r_num = start_r3 + idx
        ws3.row_dimensions[r_num].height = 19
        sym = str(r["Symbol"])
        in_p = str(r["In_Portfolio"])
        url = f"https://www.google.com/finance/quote/{sym}:NSE"

        ws3.cell(row=r_num, column=1, value=sym)
        ws3.cell(row=r_num, column=2, value=str(r["Company_Name"]))
        ws3.cell(row=r_num, column=3, value=round(float(r["LTP"]), 2))
        ws3.cell(row=r_num, column=4, value=round(float(r["DMA_200"]), 2))
        ws3.cell(row=r_num, column=5, value=f"=(C{r_num}-D{r_num})/D{r_num}")
        ws3.cell(row=r_num, column=6, value=f'=IF(C{r_num}>D{r_num},"YES","NO")')
        ws3.cell(row=r_num, column=7, value=int(r["RS_Rank"]))
        ws3.cell(row=r_num, column=8, value=in_p)
        c_u3 = ws3.cell(row=r_num, column=9, value=f"NSE: {sym}")
        c_u3.hyperlink = url

        row_fill = fill_alt if r_num % 2 == 0 else None
        for c_idx in range(1, 10):
            c = ws3.cell(row=r_num, column=c_idx)
            c.font = font_bold if c_idx == 1 else (font_link if c_idx == 9 else font_data)
            c.border = cell_border
            if c_idx == 8 and in_p == "YES":
                c.fill = fill_add
            elif row_fill:
                c.fill = row_fill
            c.alignment = align_left if c_idx in (1, 2, 9) else (align_center if c_idx in (6, 8) else align_right)

        ws3.cell(row=r_num, column=3).number_format = "₹#,##0.00"
        ws3.cell(row=r_num, column=4).number_format = "₹#,##0.00"
        ws3.cell(row=r_num, column=5).number_format = "+0.00%;-0.00%;0.00%"
        ws3.cell(row=r_num, column=7).number_format = "#,##0"

    end_r3 = start_r3 + len(screener_df) - 1
    tot_r3 = end_r3 + 1
    ws3.row_dimensions[tot_r3].height = 22
    ws3.cell(row=tot_r3, column=1, value="Average / Count")
    ws3.cell(row=tot_r3, column=2, value="All 50 Nifty 50 Constituents")
    ws3.cell(row=tot_r3, column=3, value=f"=AVERAGE(C{start_r3}:C{end_r3})")
    ws3.cell(row=tot_r3, column=4, value=f"=AVERAGE(D{start_r3}:D{end_r3})")
    ws3.cell(row=tot_r3, column=5, value=f"=AVERAGE(E{start_r3}:E{end_r3})")
    ws3.cell(row=tot_r3, column=6, value="YES")
    ws3.cell(row=tot_r3, column=7, value=f"=COUNT(G{start_r3}:G{end_r3})")
    ws3.cell(row=tot_r3, column=8, value="YES")
    c_tot3_u = ws3.cell(row=tot_r3, column=9, value="NSE Nifty 50 Tracker")
    c_tot3_u.hyperlink = "https://www.nseindia.com/index-tracker/NIFTY%2050"

    for c_idx in range(1, 10):
        c = ws3.cell(row=tot_r3, column=c_idx)
        c.font = font_bold if c_idx != 9 else font_link
        c.border = total_border
        c.alignment = align_left if c_idx in (1, 2, 9) else (align_center if c_idx in (6, 8) else align_right)
    ws3.cell(row=tot_r3, column=3).number_format = "₹#,##0.00"
    ws3.cell(row=tot_r3, column=4).number_format = "₹#,##0.00"
    ws3.cell(row=tot_r3, column=5).number_format = "+0.00%;-0.00%;0.00%"
    ws3.cell(row=tot_r3, column=7).number_format = "#,##0"

    chart3 = BarChart()
    chart3.type = "col"
    chart3.style = 10
    chart3.title = "Top 15 Nifty 50 Stocks by % Distance Above 200 DMA"
    chart3.y_axis.title = "% Distance Above 200 DMA"
    chart3.x_axis.title = "Symbol"
    chart3.width = 22
    chart3.height = 10
    data_ref3 = Reference(ws3, min_col=5, min_row=1, max_col=5, max_row=start_r3 + 14)
    cats_ref3 = Reference(ws3, min_col=1, min_row=start_r3, max_row=start_r3 + 14)
    chart3.add_data(data_ref3, titles_from_data=True)
    chart3.set_categories(cats_ref3)
    chart3.legend = None
    ws3.add_chart(chart3, f"A{tot_r3 + 3}")

    ws3.freeze_panes = "A2"
    auto_fit_sheet_columns(ws3, skip_rows=(1,))

    wb.save(excel_path)


def print_console_table(rebalance_df, summary, excel_path):
    print("\n" + "=" * 124)
    print(f" NIFTY 50 MONTHLY REBALANCE REPORT — {summary['as_of_date']} (READ FROM & SAVED TO: {os.path.basename(excel_path)})")
    print("=" * 124)
    print(f" Total Portfolio Capital : ₹{summary['total_capital']:,.2f} (Invested: ₹{summary['total_invested']:,.2f} | Residual Cash: ₹{summary['new_cash_balance']:,.2f})")
    print(f" Nifty 50 Market Breadth : {summary['num_above_200dma']} / {summary['total_universe']} stocks > 200 DMA ({summary['breadth_pct']:.1f}%)")
    print(f" Active Market Regime    : {summary['regime']}")
    print(f" Monthly Actions         : {summary['num_retained']} RETAIN | {summary['num_added']} ADD (Buy) | {summary['num_removed']} REMOVE (Sell)")
    print(f" Target Weightage        : {summary['per_stock_wt_pct']:.2f}% per equity stock | {summary['liquid_wt_pct']:.2f}% Liquid ETF")
    print("-" * 124)
    print(
        f"{'Symbol':<12} | {'Action':<14} | {'Rank':>4} | {'LTP (₹)':>10} | {'200 DMA':>10} | "
        f"{'% vs DMA':>9} | {'Prev Sh':>7} | {'Weight':>7} | {'Target (₹)':>12} | {'New Sh':>6} | {'Trade (+/-)':>11}"
    )
    print("-" * 124)
    for _, r in rebalance_df.iterrows():
        rk_str = f"#{int(r['RS_Rank'])}" if 0 < int(r["RS_Rank"]) < 50 else "-"
        print(
            f"{r['Symbol']:<12} | {r['Action']:<14} | {rk_str:>4} | {r['LTP']:>10,.2f} | {r['DMA_200']:>10,.2f} | "
            f"{r['Dist_Pct']:>+8.2f}% | {int(r['Prev_Shares']):>7,d} | {r['Target_Weight']*100:>6.2f}% | "
            f"₹{r['Target_Value']:>11,.2f} | {int(r['Target_Shares']):>6,d} | {int(r['Trade_Shares']):>+11,d}"
        )
    print("=" * 124 + "\n")


def main():
    parser = argparse.ArgumentParser(
        description="Nifty 50 Monthly Portfolio Rebalancer (Reads & Updates Previous Month's Excel File Automatically)"
    )
    parser.add_argument(
        "--excel",
        type=str,
        default="Nifty50_Portfolio_Tracker.xlsx",
        help="Path to the monthly Excel workbook (reads previous month's holdings & LTP/200DMA, then writes updated sheets)",
    )
    parser.add_argument("--capital", type=float, default=None, help="Optional override for total portfolio capital in INR")
    parser.add_argument("--top-n", type=int, default=10, help="Number of top Relative Strength stocks to hold (default: 10)")
    parser.add_argument("--keep-rank", type=int, default=15, help="Rank retention buffer to prevent churn (default: 15)")
    parser.add_argument("--breadth-cutoff", type=float, default=0.50, help="Market breadth cutoff for 50%% Liquid ETF regime (default: 0.50)")
    parser.add_argument("--live", action="store_true", help="Fetch latest live NSE prices & 200 DMAs via yfinance before rebalancing")
    args = parser.parse_args()

    excel_path = args.excel

    # 1. Read previous month's holdings & screener prices from the Excel file if it exists
    prev_holdings, prev_liquid_sh, prev_cash, screener_df = load_previous_month_excel(excel_path)

    if screener_df is None or screener_df.empty:
        screener_df = pd.DataFrame(DEFAULT_NIFTY50_DATA, columns=["Symbol", "Company_Name", "LTP", "DMA_200"])

    if not prev_holdings:
        prev_holdings = {sym: sh for sym, sh, _ in INITIAL_OCT2026_HOLDINGS}
        prev_cash = INITIAL_OCT2026_CASH

    # 2. Optionally refresh live prices via yfinance if --live is passed
    if args.live:
        screener_df, ok = try_fetch_live_yfinance(screener_df)
        if ok:
            print("[INFO] Successfully updated live NSE LTP & 200 DMA values via yfinance.")
        else:
            print("[INFO] Using LTP & 200 DMA values from the Excel workbook.")

    # 3. Compute Monthly Rebalance (Points 1, 2, 3, and 4)
    rebalance_df, updated_screener_df, summary = run_rebalance_calculation(
        screener_df=screener_df,
        prev_holdings=prev_holdings,
        prev_liquid_shares=prev_liquid_sh,
        prev_cash=prev_cash,
        override_capital=args.capital,
        top_n=args.top_n,
        keep_rank=args.keep_rank,
        breadth_threshold=args.breadth_cutoff,
        defensive_equity_wt=0.50,
    )

    # 4. Write updated Excel workbook (overwriting `excel_path` so next month reads the new state)
    write_rebalance_excel(excel_path, rebalance_df, updated_screener_df, summary)
    print_console_table(rebalance_df, summary, excel_path)


if __name__ == "__main__":
    main()