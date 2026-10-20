import json
import math
import re
import uuid
from datetime import datetime
from io import BytesIO

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Trade Analyzer", layout="wide", initial_sidebar_state="collapsed")

# ──────────────────────────────────────────────────────────────────────────────
# Design tokens
# ──────────────────────────────────────────────────────────────────────────────
INK = "#111827"
MUTED = "#6b7280"
LINE = "#e5e7eb"
SURFACE = "#ffffff"
PAGE = "#f7f8fa"
ACCENT = "#2563eb"
GOOD = "#059669"
BAD = "#dc2626"
WARN = "#d97706"

CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"], [data-testid="stAppViewContainer"] {{
  font-family: 'Inter', system-ui, -apple-system, sans-serif;
  color: {INK};
}}
[data-testid="stAppViewContainer"] {{ background: {PAGE}; }}
[data-testid="stHeader"], [data-testid="stSidebar"], #MainMenu, footer {{ display: none !important; }}

.main .block-container {{
  max-width: 1180px;
  padding: 2.5rem 1.5rem 4rem;
}}

h1 {{ font-size: 1.6rem !important; font-weight: 700 !important; letter-spacing: -0.02em; padding: 0 !important; }}
h2 {{ font-size: 1.15rem !important; font-weight: 600 !important; letter-spacing: -0.01em; padding: 0 !important; margin-top: 1.5rem; }}
h3 {{ font-size: 1rem !important; font-weight: 600 !important; }}
h4 {{ font-size: .9rem !important; font-weight: 600 !important; color: {MUTED}; }}
hr {{ border: none; border-top: 1px solid {LINE}; margin: 1.5rem 0; }}

[data-testid="stMetric"] {{
  background: {SURFACE};
  border: 1px solid {LINE};
  border-radius: 8px;
  padding: .85rem 1rem;
}}
[data-testid="stMetricLabel"] p {{ font-size: .78rem; color: {MUTED}; font-weight: 500; }}
[data-testid="stMetricValue"] {{ font-size: 1.35rem; font-weight: 600; font-variant-numeric: tabular-nums; }}
[data-testid="stMetricDelta"] {{ font-size: .75rem; }}

.stButton > button, [data-testid="stDownloadButton"] > button, [data-testid="stFormSubmitButton"] > button {{
  background: {INK};
  color: #fff;
  border: 1px solid {INK};
  border-radius: 6px;
  font-weight: 500;
  font-size: .875rem;
  padding: .45rem 1rem;
  box-shadow: none;
}}
.stButton > button:hover, [data-testid="stDownloadButton"] > button:hover, [data-testid="stFormSubmitButton"] > button:hover {{
  background: #374151; border-color: #374151; color: #fff;
}}
[data-testid="stDownloadButton"] > button {{
  background: {SURFACE}; color: {INK}; border: 1px solid {LINE};
}}
[data-testid="stDownloadButton"] > button:hover {{
  background: {PAGE}; color: {INK}; border-color: #d1d5db;
}}

[data-testid="stFileUploaderDropzone"] {{
  background: {SURFACE};
  border: 1px dashed #cbd5e1;
  border-radius: 8px;
}}
.stTextInput input, .stTextArea textarea, [data-baseweb="select"] > div {{
  border-radius: 6px !important;
  border-color: {LINE} !important;
  background: {SURFACE} !important;
}}

[data-testid="stTabs"] [role="tablist"] {{ border-bottom: 1px solid {LINE}; gap: .25rem; }}
[data-testid="stTabs"] [role="tab"] {{ font-size: .875rem; font-weight: 500; color: {MUTED}; padding: .5rem .9rem; }}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{ color: {INK}; font-weight: 600; }}

[data-testid="stExpander"] {{
  background: {SURFACE};
  border: 1px solid {LINE};
  border-radius: 8px;
  box-shadow: none;
}}
[data-testid="stExpander"] summary {{ font-weight: 500; }}

[data-testid="stDataFrame"] {{ border: 1px solid {LINE}; border-radius: 8px; overflow: hidden; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ border-color: {LINE}; border-radius: 8px; background: {SURFACE}; }}
[data-testid="stAlert"] {{ border-radius: 8px; }}

.status {{
  background: {SURFACE};
  border: 1px solid {LINE};
  border-left: 4px solid var(--c);
  border-radius: 8px;
  padding: .9rem 1.1rem;
  margin: .25rem 0 1rem;
}}
.status .head {{ font-weight: 600; color: var(--c); }}
.status .sub {{ color: {MUTED}; font-size: .875rem; margin-top: .15rem; }}
.status .sub b {{ color: {INK}; font-weight: 500; }}
</style>
"""


def inject_css():
    st.markdown(CSS, unsafe_allow_html=True)


inject_css()

if "ip_history" not in st.session_state:
    st.session_state.ip_history = []


# ──────────────────────────────────────────────────────────────────────────────
# Analysis
# ──────────────────────────────────────────────────────────────────────────────
def analyze_trades(uploaded_file, scalping_threshold_seconds: int = 180):
    """Analyze trading positions from an MT5 Excel report."""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        start_idx = df.index[df[0].astype(str).str.contains('Positions', case=False, na=False)].tolist()
        end_idx = df.index[df[0].astype(str).str.contains('Orders', case=False, na=False)].tolist()

        if not start_idx:
            st.error("Could not find a 'Positions' section in the file.")
            return None

        start = start_idx[0] + 1
        end = end_idx[0] if end_idx else len(df)
        positions_raw = df.iloc[start:end].dropna(how='all')
        positions_raw.columns = positions_raw.iloc[0]
        positions_df = positions_raw[1:].reset_index(drop=True)

        positions_df = positions_df.rename(columns={'Time': 'Open Time', 'Price': 'Open Price'})
        positions_df.columns.values[8] = 'Close Time'
        positions_df.columns.values[9] = 'Close Price'

        for col in ('Open Time', 'Close Time'):
            positions_df[col] = pd.to_datetime(
                positions_df[col], format='%Y.%m.%d %H:%M:%S.%f', errors='coerce'
            ).fillna(pd.to_datetime(positions_df[col], format='%Y.%m.%d %H:%M:%S', errors='coerce'))

        positions_df['Profit'] = pd.to_numeric(positions_df['Profit'], errors='coerce')
        positions_df['Volume'] = pd.to_numeric(positions_df.get('Volume', 0), errors='coerce').fillna(0)

        positions_df['Hold_Time'] = positions_df['Close Time'] - positions_df['Open Time']
        total_volume = positions_df['Volume'].sum()

        scalping_tdelta = pd.Timedelta(seconds=int(scalping_threshold_seconds))
        scalping_df = positions_df[positions_df['Hold_Time'] <= scalping_tdelta]

        # Reversals: opposite side, same symbol, opened within 20s of the previous close
        positions_df = positions_df.sort_values(by='Open Time').reset_index(drop=True)
        positions_df['Reversal'] = False
        for i in range(1, len(positions_df)):
            prev_close = positions_df.loc[i - 1, 'Close Time']
            curr_open = positions_df.loc[i, 'Open Time']
            prev_type = str(positions_df.loc[i - 1, 'Type']).strip().lower()
            curr_type = str(positions_df.loc[i, 'Type']).strip().lower()
            prev_symbol = str(positions_df.loc[i - 1, 'Symbol']).strip().upper()
            curr_symbol = str(positions_df.loc[i, 'Symbol']).strip().upper()

            if pd.notnull(prev_close) and pd.notnull(curr_open) and prev_symbol == curr_symbol:
                time_diff = abs((curr_open - prev_close).total_seconds())
                if time_diff <= 20 and {prev_type, curr_type} == {'buy', 'sell'}:
                    positions_df.loc[i, 'Reversal'] = True

        reversal_df = positions_df[positions_df['Reversal']]
        reversal_count = len(reversal_df)
        reversal_profit = reversal_df['Profit'].sum()

        # Bursts: 2+ consecutive opens within 2 seconds of each other
        positions_df['Burst'] = False
        burst_groups, current_group = [], [0]
        for i in range(1, len(positions_df)):
            prev_open = positions_df.loc[i - 1, 'Open Time']
            curr_open = positions_df.loc[i, 'Open Time']
            if pd.notnull(prev_open) and pd.notnull(curr_open):
                if abs((curr_open - prev_open).total_seconds()) <= 2:
                    current_group.append(i)
                else:
                    if len(current_group) >= 2:
                        burst_groups.append(current_group)
                    current_group = [i]
        if len(current_group) >= 2:
            burst_groups.append(current_group)
        for group in burst_groups:
            positions_df.loc[group, 'Burst'] = True

        burst_df = positions_df[positions_df['Burst']]
        burst_count = len(burst_df)
        burst_profit = burst_df['Profit'].sum()

        total_positions = len(positions_df)
        total_profit = positions_df['Profit'].sum()
        scalping_count = len(scalping_df)
        scalping_profit = scalping_df['Profit'].sum()

        def pct(n, d):
            return (n / d * 100) if d else 0

        scalping_percentage = pct(scalping_count, total_positions)
        reversal_percentage = pct(reversal_count, total_positions)
        burst_percentage = pct(burst_count, total_positions)
        scalping_profit_percentage = pct(scalping_profit, total_profit) if total_profit != 0 else 0
        reversal_profit_percentage = pct(reversal_profit, total_profit) if total_profit != 0 else 0
        burst_profit_percentage = pct(burst_profit, total_profit) if total_profit != 0 else 0

        avg_hold_time = positions_df['Hold_Time'].mean()
        avg_scalping_hold_time = scalping_df['Hold_Time'].mean() if len(scalping_df) > 0 else pd.Timedelta(0)

        wins = positions_df[positions_df['Profit'] > 0]
        losses = positions_df[positions_df['Profit'] < 0]
        win_rate = pct(len(wins), total_positions)
        avg_win = wins['Profit'].mean() if len(wins) > 0 else 0.0
        avg_loss = losses['Profit'].mean() if len(losses) > 0 else 0.0
        gross_profit = wins['Profit'].sum()
        gross_loss = losses['Profit'].sum()
        profit_factor = gross_profit / abs(gross_loss) if gross_loss != 0 else 0.0
        max_win = positions_df['Profit'].max()
        max_loss = positions_df['Profit'].min()
        expectancy = (win_rate / 100 * avg_win) + ((1 - win_rate / 100) * avg_loss)

        avg_winner_hold = wins['Hold_Time'].mean() if len(wins) > 0 else pd.Timedelta(0)
        avg_loser_hold = losses['Hold_Time'].mean() if len(losses) > 0 else pd.Timedelta(0)
        median_hold = positions_df['Hold_Time'].median()

        eq_df = positions_df.sort_values('Close Time').reset_index(drop=True)
        equity_curve = eq_df['Profit'].cumsum()
        drawdown_series = equity_curve - equity_curve.cummax()
        max_drawdown = drawdown_series.min()

        max_streak = streak = 0
        for p in eq_df['Profit']:
            if p < 0:
                streak += 1
                max_streak = max(max_streak, streak)
            else:
                streak = 0
        worst_loss_streak = max_streak

        daily_pnl = eq_df.groupby(eq_df['Close Time'].dt.date)['Profit'].sum()
        best_day = daily_pnl.max()
        worst_day = daily_pnl.min()

        buy_df = positions_df[positions_df['Type'].astype(str).str.lower() == 'buy']
        sell_df = positions_df[positions_df['Type'].astype(str).str.lower() == 'sell']

        symbol_counts = positions_df['Symbol'].value_counts()
        top_symbol = symbol_counts.index[0] if len(symbol_counts) > 0 else "N/A"
        symbol_concentration = symbol_counts.iloc[0] / total_positions * 100 if total_positions > 0 else 0.0

        hourly = positions_df.groupby(positions_df['Open Time'].dt.hour)['Profit'].agg(['sum', 'count', 'mean'])
        hourly.index.name = 'Hour'
        dow_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        daily_dow = positions_df.groupby(positions_df['Open Time'].dt.day_name())['Profit'].agg(['sum', 'count', 'mean'])
        daily_dow = daily_dow.reindex([d for d in dow_order if d in daily_dow.index])

        martingale_flags = sum(
            1 for i in range(1, len(eq_df))
            if eq_df.loc[i - 1, 'Profit'] < 0 and eq_df.loc[i, 'Volume'] > eq_df.loc[i - 1, 'Volume']
        )
        martingale_score = martingale_flags / (total_positions - 1) * 100 if total_positions > 1 else 0.0

        toxic_count = scalping_count + reversal_count + burst_count
        toxic_percentage = pct(toxic_count, total_positions)

        avg_hold_minutes = avg_hold_time.total_seconds() / 60 if pd.notna(avg_hold_time) else 0
        if avg_hold_minutes < 5 and scalping_percentage > 50:
            flow_classification = "High-risk scalper flow"
        elif martingale_score > 30:
            flow_classification = "Martingale/averaging flow"
        elif avg_hold_minutes < 60:
            flow_classification = "Intraday active flow"
        else:
            flow_classification = "Swing/positional flow"

        if profit_factor < 0.8 and max_drawdown < -5000 and toxic_percentage > 50:
            dealer_decision = "B-Book – high confidence"
        elif profit_factor < 1.0:
            dealer_decision = "B-Book – moderate confidence"
        elif profit_factor > 1.5 and max_drawdown > -3000:
            dealer_decision = "A-Book candidate"
        else:
            dealer_decision = "Review required"

        positions_df['Scalping'] = positions_df['Hold_Time'] <= scalping_tdelta

        THRESHOLD_PCT = 0.70

        winners_sorted = positions_df[positions_df['Profit'] > 0].sort_values('Profit', ascending=False)
        gross_profit_sum = winners_sorted['Profit'].sum()
        top_profit_contributors_df = pd.DataFrame()
        top_profit_count = 0
        top_profit_pct_of_count = 0.0
        top_profit_pct_of_gross = 0.0
        if gross_profit_sum > 0 and len(winners_sorted) > 0:
            winners_sorted = winners_sorted.copy()
            winners_sorted['Pct_of_Gross_Profit'] = (winners_sorted['Profit'] / gross_profit_sum * 100).round(2)
            winners_sorted['Cumulative_Profit'] = winners_sorted['Profit'].cumsum().round(2)
            winners_sorted['Cumulative_Pct'] = (winners_sorted['Cumulative_Profit'] / gross_profit_sum * 100).round(2)
            cutoff_mask = winners_sorted['Cumulative_Profit'] >= gross_profit_sum * THRESHOLD_PCT
            if cutoff_mask.any():
                cutoff_pos = winners_sorted.index.get_loc(cutoff_mask.idxmax())
                top_profit_contributors_df = winners_sorted.iloc[:cutoff_pos + 1].copy()
                top_profit_count = len(top_profit_contributors_df)
                top_profit_pct_of_count = pct(top_profit_count, total_positions)
                top_profit_pct_of_gross = top_profit_contributors_df['Profit'].sum() / gross_profit_sum * 100

        losers_sorted = positions_df[positions_df['Profit'] < 0].sort_values('Profit', ascending=True)
        gross_loss_abs = abs(losers_sorted['Profit'].sum())
        top_loss_contributors_df = pd.DataFrame()
        top_loss_count = 0
        top_loss_pct_of_count = 0.0
        top_loss_pct_of_gross = 0.0
        if gross_loss_abs > 0 and len(losers_sorted) > 0:
            losers_sorted = losers_sorted.copy()
            losers_sorted['Pct_of_Gross_Loss'] = (losers_sorted['Profit'].abs() / gross_loss_abs * 100).round(2)
            losers_sorted['Cumulative_Loss'] = losers_sorted['Profit'].cumsum().round(2)
            losers_sorted['Cumulative_Pct'] = (losers_sorted['Cumulative_Loss'].abs() / gross_loss_abs * 100).round(2)
            cutoff_mask = losers_sorted['Cumulative_Loss'] <= -gross_loss_abs * THRESHOLD_PCT
            if cutoff_mask.any():
                cutoff_pos = losers_sorted.index.get_loc(cutoff_mask.idxmax())
                top_loss_contributors_df = losers_sorted.iloc[:cutoff_pos + 1].copy()
                top_loss_count = len(top_loss_contributors_df)
                top_loss_pct_of_count = pct(top_loss_count, total_positions)
                top_loss_pct_of_gross = top_loss_contributors_df['Profit'].abs().sum() / gross_loss_abs * 100

        max_win_row = positions_df.loc[[positions_df['Profit'].idxmax()]] if total_positions > 0 else pd.DataFrame()
        max_loss_row = positions_df.loc[[positions_df['Profit'].idxmin()]] if total_positions > 0 else pd.DataFrame()

        return {
            "total_positions": total_positions,
            "total_profit": total_profit,
            "total_volume": total_volume,
            "scalping_count": scalping_count,
            "scalping_profit": scalping_profit,
            "scalping_percentage": scalping_percentage,
            "scalping_profit_percentage": scalping_profit_percentage,
            "reversal_count": reversal_count,
            "reversal_profit": reversal_profit,
            "reversal_percentage": reversal_percentage,
            "reversal_profit_percentage": reversal_profit_percentage,
            "burst_count": burst_count,
            "burst_profit": burst_profit,
            "burst_percentage": burst_percentage,
            "burst_profit_percentage": burst_profit_percentage,
            "avg_hold_time": avg_hold_time,
            "avg_scalping_hold_time": avg_scalping_hold_time,
            "scalping_df": scalping_df,
            "reversal_df": reversal_df,
            "burst_df": burst_df,
            "all_positions_df": positions_df,
            "win_rate": win_rate,
            "avg_win": avg_win,
            "avg_loss": avg_loss,
            "profit_factor": profit_factor,
            "max_win": max_win,
            "max_loss": max_loss,
            "expectancy": expectancy,
            "avg_winner_hold": avg_winner_hold,
            "avg_loser_hold": avg_loser_hold,
            "median_hold": median_hold,
            "eq_df": eq_df,
            "equity_curve": equity_curve,
            "drawdown_series": drawdown_series,
            "max_drawdown": max_drawdown,
            "worst_loss_streak": worst_loss_streak,
            "daily_pnl": daily_pnl,
            "best_day": best_day,
            "worst_day": worst_day,
            "buy_count": len(buy_df),
            "buy_net": buy_df['Profit'].sum(),
            "sell_count": len(sell_df),
            "sell_net": sell_df['Profit'].sum(),
            "top_symbol": top_symbol,
            "symbol_concentration": symbol_concentration,
            "hourly": hourly,
            "daily_dow": daily_dow,
            "martingale_score": martingale_score,
            "toxic_percentage": toxic_percentage,
            "flow_classification": flow_classification,
            "dealer_decision": dealer_decision,
            "top_profit_contributors_df": top_profit_contributors_df,
            "top_profit_count": top_profit_count,
            "top_profit_pct_of_count": top_profit_pct_of_count,
            "top_profit_pct_of_gross": top_profit_pct_of_gross,
            "top_loss_contributors_df": top_loss_contributors_df,
            "top_loss_count": top_loss_count,
            "top_loss_pct_of_count": top_loss_pct_of_count,
            "top_loss_pct_of_gross": top_loss_pct_of_gross,
            "max_win_row": max_win_row,
            "max_loss_row": max_loss_row,
        }

    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None


@st.cache_data(show_spinner=False)
def _analyze_bytes(file_bytes: bytes, filename: str, scalping_threshold_seconds: int = 180):
    return analyze_trades(BytesIO(file_bytes), scalping_threshold_seconds)


# ──────────────────────────────────────────────────────────────────────────────
# Exports and reports
# ──────────────────────────────────────────────────────────────────────────────
def _build_overview_excel(result: dict, account_no: str) -> BytesIO:
    buf = BytesIO()
    pos = result['all_positions_df'].copy()

    def _yn(col):
        return pos[col].fillna(False).map({True: 'Y', False: ''}) if col in pos.columns else ''

    max_w, max_l = result.get('max_win'), result.get('max_loss')
    top_p_idx = set(result['top_profit_contributors_df'].index)
    top_l_idx = set(result['top_loss_contributors_df'].index)

    pos['Is_Scalping'] = _yn('Scalping')
    pos['Is_Reversal'] = _yn('Reversal')
    pos['Is_Burst'] = _yn('Burst')
    pos['Is_Max_Win'] = pos['Profit'].apply(lambda p: 'Y' if p == max_w else '')
    pos['Is_Max_Loss'] = pos['Profit'].apply(lambda p: 'Y' if p == max_l else '')
    pos['Is_Top_Profit_Contributor'] = pos.index.map(lambda i: 'Y' if i in top_p_idx else '')
    pos['Is_Top_Loss_Contributor'] = pos.index.map(lambda i: 'Y' if i in top_l_idx else '')

    base_cols = [c for c in ['Symbol', 'Type', 'Open Time', 'Close Time', 'Volume', 'Profit', 'Hold_Time'] if c in pos.columns]
    flag_cols = ['Is_Scalping', 'Is_Reversal', 'Is_Burst', 'Is_Max_Win', 'Is_Max_Loss',
                 'Is_Top_Profit_Contributor', 'Is_Top_Loss_Contributor']

    summary_rows = [
        ('Account', account_no),
        ('Total Trades', result['total_positions']),
        ('Net Profit', f"${result['total_profit']:.2f}"),
        ('Win Rate %', f"{result['win_rate']:.2f}"),
        ('Profit Factor', f"{result['profit_factor']:.2f}"),
        ('Max Win', f"${result['max_win']:.2f}"),
        ('Max Loss', f"${result['max_loss']:.2f}"),
        ('Max Drawdown', f"${result['max_drawdown']:.2f}"),
        ('Worst Loss Streak', result['worst_loss_streak']),
        ('Scalping Trades', result['scalping_count']),
        ('Reversal Trades', result['reversal_count']),
        ('Burst Trades', result['burst_count']),
        ('Toxic %', f"{result['toxic_percentage']:.2f}"),
        ('Top profit contributors (>=70% of gross profit)', result['top_profit_count']),
        ('  % of total trade count', f"{result['top_profit_pct_of_count']:.2f}%"),
        ('  % of gross profit explained', f"{result['top_profit_pct_of_gross']:.2f}%"),
        ('Top loss contributors (>=70% of gross loss)', result['top_loss_count']),
        ('  % of total trade count', f"{result['top_loss_pct_of_count']:.2f}%"),
        ('  % of gross loss explained', f"{result['top_loss_pct_of_gross']:.2f}%"),
        ('Dealer Decision', result['dealer_decision']),
        ('Flow Classification', result['flow_classification']),
    ]

    extra_sheets = [
        ('Top Profit (>=70%)', result['top_profit_contributors_df']),
        ('Top Loss (>=70%)', result['top_loss_contributors_df']),
        ('Max Win Trade', result['max_win_row']),
        ('Max Loss Trade', result['max_loss_row']),
        ('Scalping Trades', result['scalping_df']),
        ('Reversal Trades', result['reversal_df']),
        ('Burst Trades', result['burst_df']),
    ]

    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        pd.DataFrame(summary_rows, columns=['Metric', 'Value']).to_excel(w, index=False, sheet_name='Summary')
        pos[base_cols + flag_cols].to_excel(w, index=False, sheet_name='All Trades (Flagged)')
        for name, sheet_df in extra_sheets:
            if not sheet_df.empty:
                sheet_df.to_excel(w, index=False, sheet_name=name)

    buf.seek(0)
    return buf


def _build_comparison_df(results):
    rows = []
    for fname, acc, r in results:
        if r is None:
            continue
        rows.append({
            "Account": acc,
            "File": fname,
            "Trades": r["total_positions"],
            "Net Profit": float(r["total_profit"]),
            "Win Rate %": round(r["win_rate"], 2),
            "Profit Factor": round(r["profit_factor"], 2),
            "Expectancy": round(r["expectancy"], 2),
            "Max Drawdown": float(r["max_drawdown"]),
            "Loss Streak": r["worst_loss_streak"],
            "Top Symbol": r["top_symbol"],
            "Concentration %": round(r["symbol_concentration"], 2),
            "Toxic %": round(r["toxic_percentage"], 2),
            "Flow": r["flow_classification"],
            "Dealer Decision": r["dealer_decision"],
        })
    return pd.DataFrame(rows)


def _trading_pattern(r: dict) -> str:
    notes = []
    if r["scalping_percentage"] > 30:
        notes.append("Scalping")
    if r["reversal_count"] > 0 and r["reversal_percentage"] > 0:
        notes.append("Hedging/Reversal")
    if r["burst_count"] > 5:
        notes.append("Burst/High-Frequency")
    return ", ".join(notes) if notes else "Normal Trading Flow"


def _build_client_report(account_no: str, r: dict, scalping_choice: int) -> str:
    total = r['total_positions']
    toxic_br = (r['reversal_count'] + r['burst_count']) / total * 100 if total > 0 else 0.0
    return (
        f"Account: {account_no}\n"
        f"Total Trades: {r['total_positions']}\n"
        f"Total Profit: ${r['total_profit']:.2f}\n"
        f"Scalping Trades: {r['scalping_count']} ({r['scalping_percentage']:.1f}% of total)\n"
        f"Scalping Profit: ${r['scalping_profit']:.2f} ({r['scalping_profit_percentage']:.1f}% of total profit)\n"
        f"Reversal Trades: {r['reversal_count']} ({r['reversal_percentage']:.1f}% of total)\n"
        f"Reversal Profit: ${r['reversal_profit']:.2f} ({r['reversal_profit_percentage']:.1f}% of total profit)\n"
        f"Burst Trades: {r['burst_count']} [{r['burst_percentage']:.1f}% of total]\n"
        f"Burst Profit: ${r['burst_profit']:.2f} [{r['burst_profit_percentage']:.1f}% of total profit]\n"
        f"Toxic Trading %: {r['toxic_percentage']:.1f}%\n"
        f"Toxic Trading (b+r) : {toxic_br:.1f}%\n"
        f"Trading Pattern: {_trading_pattern(r)}\n"
        f"Scalping Threshold Used: {scalping_choice} seconds"
    )


# ──────────────────────────────────────────────────────────────────────────────
# UI helpers
# ──────────────────────────────────────────────────────────────────────────────
def fmt_td(td):
    if pd.isna(td):
        return "N/A"
    total_s = int(td.total_seconds())
    h, rem = divmod(abs(total_s), 3600)
    m, s = divmod(rem, 60)
    return f"{h}h {m}m {s}s" if h else f"{m}m {s}s"


def metric_row(items):
    """items: list of (label, value) or (label, value, delta)."""
    cols = st.columns(len(items))
    for col, item in zip(cols, items):
        if len(item) == 3:
            col.metric(item[0], item[1], delta=item[2], delta_color="off")
        else:
            col.metric(item[0], item[1])


def status_banner(headline: str, sub: str, color: str):
    st.markdown(
        f"<div class='status' style='--c:{color}'>"
        f"<div class='head'>{headline}</div><div class='sub'>{sub}</div></div>",
        unsafe_allow_html=True,
    )


def dealer_color(decision: str) -> str:
    if "B-Book" in decision:
        return BAD
    if "A-Book" in decision:
        return GOOD
    return WARN


def style_fig(fig, height):
    fig.update_layout(
        template="simple_white",
        height=height,
        margin=dict(t=40, b=20, l=0, r=0),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", size=12, color=INK),
        title=dict(font=dict(size=14)),
    )
    return fig


def risk_flags(r: dict):
    flags = []
    if r['symbol_concentration'] > 80:
        flags.append(f"High symbol concentration: {r['symbol_concentration']:.1f}% of trades on **{r['top_symbol']}**")
    if r['max_drawdown'] < -10000:
        flags.append(f"Severe drawdown: **${r['max_drawdown']:.2f}**")
    elif r['max_drawdown'] < -3000:
        flags.append(f"Elevated drawdown: **${r['max_drawdown']:.2f}**")
    if r['profit_factor'] < 1.0:
        flags.append("Profit factor below 1.0 (net losing strategy)")
    if r['martingale_score'] > 30:
        flags.append(f"Martingale/averaging pattern: {r['martingale_score']:.1f}%")
    if r['avg_loss'] != 0 and r['avg_win'] > 0 and abs(r['avg_loss']) > r['avg_win'] * 3:
        flags.append(f"Loss asymmetry: average loss is {abs(r['avg_loss']) / r['avg_win']:.1f}x the average win")
    if r['toxic_percentage'] > 50:
        flags.append(f"High flagged trade rate: {r['toxic_percentage']:.1f}%")
    return flags


def render_copy_button(report_text: str, key: str, show_code: bool = True, label: str = "Copy report"):
    if show_code:
        st.code(report_text, language=None)
    escaped = json.dumps(report_text)
    label_js = json.dumps(label)
    btn_id = f"copyBtn_{key}"
    components.html(
        f"""
        <button id="{btn_id}" style="width:100%;background:#fff;color:{INK};border:1px solid {LINE};border-radius:6px;padding:.5rem 1rem;font:500 14px Inter,system-ui,sans-serif;cursor:pointer;">{label}</button>
        <script>
        (function(){{
            var b=document.getElementById('{btn_id}');
            if(!b)return;
            b.addEventListener('click',async function(){{
                var t={escaped}, ok=false;
                try{{
                    if(navigator.clipboard && window.isSecureContext){{ await navigator.clipboard.writeText(t); ok=true; }}
                }}catch(e){{ ok=false; }}
                if(!ok){{
                    try{{
                        var ta=document.createElement('textarea');
                        ta.value=t; ta.setAttribute('readonly','');
                        ta.style.cssText='position:fixed;top:0;left:0;opacity:0';
                        document.body.appendChild(ta); ta.focus(); ta.select();
                        ta.setSelectionRange(0,t.length);
                        ok=document.execCommand('copy');
                        document.body.removeChild(ta);
                    }}catch(e){{ ok=false; }}
                }}
                b.textContent = ok ? 'Copied' : 'Copy failed. Select the text manually.';
                setTimeout(function(){{ b.textContent={label_js}; }}, 2000);
            }});
        }})();
        </script>
        """,
        height=46,
    )


def download_pair(df, base_name: str, key: str, sheet: str, label: str):
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(f"{label} (CSV)", df.to_csv(index=False).encode(), f"{base_name}.csv", "text/csv", key=f"{key}_csv", use_container_width=True)
    with c2:
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False, sheet_name=sheet[:31])
        buf.seek(0)
        st.download_button(
            f"{label} (Excel)", buf, f"{base_name}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key}_xl", use_container_width=True,
        )


def win_rate_of(df):
    return f"{(df['Profit'] > 0).mean() * 100:.1f}%" if len(df) > 0 else "N/A"


def mean_profit_of(df):
    return f"${df['Profit'].mean():.2f}" if len(df) > 0 else "N/A"


def pattern_section(title, df, count, pct_total, profit, pct_profit, key, account_no, name, avg_time=None):
    st.markdown(f"### {title}")
    items = [
        (f"{name} trades", count, f"{pct_total:.1f}% of total"),
        (f"{name} profit", f"${profit:.2f}", f"{pct_profit:.1f}% of profit"),
        ("Win rate", win_rate_of(df)),
        ("Avg time" if avg_time is not None else "Avg profit",
         fmt_td(avg_time) if avg_time is not None else mean_profit_of(df)),
    ]
    metric_row(items)
    if count > 0:
        download_pair(df, f"{key}_trades_{account_no}", key, name, name)


# ──────────────────────────────────────────────────────────────────────────────
# Bulk view
# ──────────────────────────────────────────────────────────────────────────────
def render_account_expander(filename: str, account_no: str, r: dict, scalping_choice: int):
    dealer = r["dealer_decision"]
    short_name = filename if len(filename) <= 40 else filename[:37] + "..."
    with st.expander(f"Account {account_no}  |  {short_name}  |  {dealer}"):
        status_banner(dealer, f"Flow: <b>{r['flow_classification']}</b>", dealer_color(dealer))
        metric_row([("Net profit", f"${r['total_profit']:.2f}"), ("Win rate %", f"{r['win_rate']:.2f}"), ("Profit factor", f"{r['profit_factor']:.2f}")])
        metric_row([("Max drawdown", f"${r['max_drawdown']:.2f}"), ("Worst loss streak", r['worst_loss_streak']), ("Symbol concentration %", f"{r['symbol_concentration']:.2f}")])
        metric_row([("Avg win", f"${r['avg_win']:.2f}"), ("Avg loss", f"${r['avg_loss']:.2f}"), ("Toxic %", f"{r['toxic_percentage']:.2f}")])

        flags = risk_flags(r)
        if flags:
            st.markdown("**Risk flags**")
            for f in flags:
                st.markdown(f"- {f}")

        st.markdown("**Client report**")
        render_copy_button(_build_client_report(account_no, r, scalping_choice), f"{account_no}_{uuid.uuid4().hex[:6]}")


def render_bulk_view(uploaded_files, scalping_choice: int):
    st.subheader(f"Bulk comparison ({len(uploaded_files)} files)")

    results, errors = [], []
    prog = st.progress(0.0)
    status = st.empty()

    for i, f in enumerate(uploaded_files):
        status.caption(f"Processing {f.name} ({i + 1}/{len(uploaded_files)})")
        try:
            r = _analyze_bytes(f.getvalue(), f.name, scalping_choice)
            if r is None:
                errors.append((f.name, "Could not parse file (no Positions section found)"))
            else:
                m = re.search(r"ReportHistory[-_ ]?(\d+)", f.name)
                results.append((f.name, m.group(1) if m else "Unknown", r))
        except Exception as exc:
            errors.append((f.name, str(exc)))
        prog.progress((i + 1) / len(uploaded_files))

    status.empty()
    prog.empty()

    for fname, err in errors:
        st.error(f"{fname}: {err}")
    if not results:
        st.warning("No files were processed successfully.")
        return

    decisions = [r["dealer_decision"] for _, _, r in results]
    metric_row([
        ("Accounts", len(results)),
        ("Aggregate net P&L", f"${sum(r['total_profit'] for _, _, r in results):.2f}"),
        ("Total trades", sum(r['total_positions'] for _, _, r in results)),
        ("Decision mix", f"{sum('A-Book' in d for d in decisions)} A-Book",
         f"{sum('B-Book – high' in d for d in decisions)} B-high, {sum('B-Book – moderate' in d for d in decisions)} B-mod, {sum('Review' in d for d in decisions)} review"),
    ])

    st.markdown("### Comparison")
    df_summary = _build_comparison_df(results)
    st.dataframe(
        df_summary,
        use_container_width=True,
        height=min(400, 60 + 35 * len(df_summary)),
        column_config={
            "Net Profit": st.column_config.NumberColumn(format="$%.2f"),
            "Max Drawdown": st.column_config.NumberColumn(format="$%.2f"),
            "Expectancy": st.column_config.NumberColumn(format="$%.2f"),
            "Win Rate %": st.column_config.NumberColumn(format="%.2f"),
            "Profit Factor": st.column_config.NumberColumn(format="%.2f"),
            "Concentration %": st.column_config.NumberColumn(format="%.2f"),
            "Toxic %": st.column_config.NumberColumn(format="%.2f"),
        },
    )
    st.download_button("Download comparison (CSV)", df_summary.to_csv(index=False).encode("utf-8"),
                       file_name="bulk_comparison.csv", mime="text/csv", key="bulk_csv_download")

    st.markdown("### Accounts")
    for fname, acc, r in results:
        render_account_expander(fname, acc, r, scalping_choice)

    st.markdown("### Client reports")
    sep = "\n" + "=" * 60 + "\n\n"
    combined = sep.join(
        f"Account {acc}\n{'-' * 60}\n" + _build_client_report(acc, r, scalping_choice)
        for _, acc, r in results
    )
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Download all reports (.txt)", combined.encode("utf-8"),
                           file_name=f"client_reports_{datetime.now():%Y%m%d_%H%M%S}.txt",
                           mime="text/plain", key="bulk_reports_download", use_container_width=True)
    with c2:
        render_copy_button(combined, f"bulk_all_{uuid.uuid4().hex[:6]}", show_code=False, label="Copy all reports")


# ──────────────────────────────────────────────────────────────────────────────
# IP lookup
# ──────────────────────────────────────────────────────────────────────────────
def get_ip_details(ip_address):
    try:
        response = requests.get(f'https://ipinfo.io/{ip_address}/json', timeout=5)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}


def add_ip_to_history(ip, details):
    entry = {"timestamp": datetime.now().strftime("%H:%M:%S"), "ip": ip, "details": details}
    st.session_state.ip_history.insert(0, entry)
    st.session_state.ip_history = st.session_state.ip_history[:10]


# ──────────────────────────────────────────────────────────────────────────────
# Page
# ──────────────────────────────────────────────────────────────────────────────
st.title("Trade Analyzer")

st.header("Trade analysis")

up_col, thr_col = st.columns([3, 1])
with up_col:
    uploaded_files = st.file_uploader(
        "MT5 Excel reports (one for a full analysis, several to compare)",
        type=["xlsx"], key="trade_file", accept_multiple_files=True,
    )
with thr_col:
    scalping_choice = st.selectbox(
        "Scalping threshold (seconds)", options=[60, 120, 180], index=2,
        help="Trades held for this long or less count as scalping.",
    )

if uploaded_files and len(uploaded_files) > 1:
    render_bulk_view(uploaded_files, scalping_choice)

elif uploaded_files:
    uploaded_file = uploaded_files[0]
    with st.spinner("Analyzing trades..."):
        result = analyze_trades(uploaded_file, scalping_threshold_seconds=scalping_choice)

    if result:
        m = re.search(r"ReportHistory[-_ ]?(\d+)", uploaded_file.name)
        account_no = m.group(1) if m else "Unknown"
        toxic_percentage = result["toxic_percentage"]

        tab_overview, tab_decision, tab_perf, tab_trades, tab_flags = st.tabs(
            ["Overview", "Decision", "Performance", "Trades and patterns", "Flagged trades"]
        )

        # ── Overview ──
        with tab_overview:
            metric_row([
                ("Net profit", f"${result['total_profit']:.2f}"),
                ("Win rate %", f"{result['win_rate']:.2f}"),
                ("Profit factor", f"{result['profit_factor']:.2f}"),
                ("Flagged trade %", f"{toxic_percentage:.2f}"),
            ])
            st.write("")
            metric_row([
                ("Average win", f"${result['avg_win']:.2f}"),
                ("Average loss", f"${result['avg_loss']:.2f}"),
                ("Max drawdown", f"${result['max_drawdown']:.2f}"),
                ("Worst loss streak", result['worst_loss_streak']),
            ])
            st.write("")
            metric_row([
                ("Avg hold", fmt_td(result['avg_hold_time'])),
                ("Median hold", fmt_td(result['median_hold'])),
                ("Top symbol", result['top_symbol']),
                ("Total volume", f"{result['total_volume']:.2f}"),
            ])

            st.markdown("### P&L concentration (70% of gross)")
            metric_row([
                ("Top profit contributors", result['top_profit_count'], f"{result['top_profit_pct_of_gross']:.1f}% of gross profit"),
                ("Share of all trades", f"{result['top_profit_pct_of_count']:.2f}%"),
                ("Top loss contributors", result['top_loss_count'], f"{result['top_loss_pct_of_gross']:.1f}% of gross loss"),
                ("Share of all trades", f"{result['top_loss_pct_of_count']:.2f}%"),
            ])


        # ── Decision ──
        with tab_decision:
            dealer = result['dealer_decision']
            status_banner(dealer, f"Flow classification: <b>{result['flow_classification']}</b>", dealer_color(dealer))
            metric_row([
                ("Expectancy per trade", f"${result['expectancy']:.2f}"),
                ("Profit factor", f"{result['profit_factor']:.2f}"),
                ("Martingale score %", f"{result['martingale_score']:.1f}"),
            ])
            st.markdown("### Risk flags")
            flags = risk_flags(result)
            if flags:
                for f in flags:
                    st.markdown(f"- {f}")
            else:
                st.success("No critical risk flags detected.")

        # ── Performance ──
        with tab_perf:
            eq_df = result['eq_df']

            fig = px.line(x=eq_df['Close Time'], y=result['equity_curve'],
                          labels={'x': 'Date', 'y': 'Cumulative P&L'}, title='Equity curve')
            fig.update_traces(line_color=ACCENT)
            st.plotly_chart(style_fig(fig, 300), use_container_width=True)

            fig = px.area(x=eq_df['Close Time'], y=result['drawdown_series'],
                          labels={'x': 'Date', 'y': 'Drawdown'}, title='Drawdown from peak',
                          color_discrete_sequence=[BAD])
            st.plotly_chart(style_fig(fig, 250), use_container_width=True)


        # ── Trades and patterns ──
        with tab_trades:
            st.markdown("### All trades")
            all_df = result['all_positions_df']
            show_cols = [c for c in ['Symbol', 'Type', 'Open Time', 'Close Time', 'Volume', 'Profit', 'Hold_Time'] if c in all_df.columns]
            st.dataframe(all_df[show_cols], use_container_width=True, height=380)

            d1, d2 = st.columns(2)
            with d1:
                st.download_button("Download trades (CSV)", all_df[show_cols].to_csv(index=False).encode(),
                                   f"overview_{account_no}.csv", "text/csv", key="ov_csv", use_container_width=True)
            with d2:
                st.download_button("Download full analysis (Excel, multi-sheet)", _build_overview_excel(result, account_no),
                                   file_name=f"overview_{account_no}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="ov_xl", use_container_width=True)

            st.markdown("---")
            pattern_section("Scalping", result['scalping_df'], result['scalping_count'], result['scalping_percentage'],
                            result['scalping_profit'], result['scalping_profit_percentage'],
                            "scalping", account_no, "Scalping", avg_time=result['avg_scalping_hold_time'] if result['scalping_count'] else pd.NaT)
            st.markdown("---")
            pattern_section("Reversals", result['reversal_df'], result['reversal_count'], result['reversal_percentage'],
                            result['reversal_profit'], result['reversal_profit_percentage'],
                            "reversal", account_no, "Reversal")
            st.markdown("---")
            pattern_section("Bursts", result['burst_df'], result['burst_count'], result['burst_percentage'],
                            result['burst_profit'], result['burst_profit_percentage'],
                            "burst", account_no, "Burst")


        # ── Flagged trades ──
        with tab_flags:
            st.caption("Three lenses applied to this account's trade history.")
            f1, f2, f3 = st.tabs(["Profit concentration", "Same-second trades", "Max win and loss"])
            all_df = result['all_positions_df']

            with f1:
                st.markdown("The top 10% of winning trades, ranked by profit. The flag triggers if they make up 70% or more of gross profit.")
                winners = all_df[all_df['Profit'] > 0]
                gross = float(winners['Profit'].sum()) if not winners.empty else 0.0
                if gross <= 0:
                    st.caption("No profitable trades to evaluate.")
                else:
                    ws = winners.sort_values('Profit', ascending=False).reset_index(drop=True)
                    top_n = max(1, math.ceil(len(ws) * 0.10))
                    top10 = ws.iloc[:top_n].copy()
                    combined = float(top10['Profit'].sum())
                    pct_gross = combined / gross * 100
                    triggered = pct_gross >= 70.0
                    top10['Pct_of_Gross_Profit'] = (top10['Profit'] / gross * 100).round(2)
                    top10['Cumulative_Pct'] = (top10['Profit'].cumsum() / gross * 100).round(2)

                    if triggered:
                        status_banner("Flag triggered", "The top 10% of winners drive 70% or more of gross profit.", BAD)
                    else:
                        status_banner("Not triggered", "Gross profit is spread across winners.", GOOD)

                    metric_row([
                        ("Top 10% winners", top_n, f"of {len(ws)} winning trades"),
                        ("Share of all trades", f"{top_n / len(all_df) * 100:.2f}%"),
                        ("Combined profit", f"${combined:.2f}"),
                        ("Share of gross profit", f"{pct_gross:.1f}%"),
                    ])
                    cols = [c for c in ['Symbol', 'Type', 'Open Time', 'Close Time', 'Volume', 'Profit', 'Pct_of_Gross_Profit', 'Cumulative_Pct'] if c in top10.columns]
                    st.dataframe(top10[cols], use_container_width=True, height=360)
                    download_pair(top10[cols], f"top10pct_winners_{account_no}", "flag_top10", "Top10pctWinners", "Download")

            with f2:
                st.markdown("Trades opened and closed within the same second (hold time of 1 second or less).")
                hold_sec = (all_df['Hold_Time'] if 'Hold_Time' in all_df.columns else all_df['Close Time'] - all_df['Open Time']).dt.total_seconds()
                ss = all_df[hold_sec.abs() <= 1].copy()
                if ss.empty:
                    st.caption("No same-second trades found.")
                else:
                    metric_row([
                        ("Flagged trades", len(ss), f"{len(ss) / len(all_df) * 100:.1f}% of total"),
                        ("Total P&L", f"${float(ss['Profit'].sum()):.2f}"),
                        ("Win rate", win_rate_of(ss)),
                        ("Avg P&L", mean_profit_of(ss)),
                    ])
                    cols = [c for c in ['Symbol', 'Type', 'Open Time', 'Close Time', 'Volume', 'Profit', 'Hold_Time'] if c in ss.columns]
                    st.dataframe(ss[cols], use_container_width=True, height=360)
                    download_pair(ss[cols], f"same_second_trades_{account_no}", "flag_samesec", "SameSecond", "Download")

            with f3:
                st.markdown("The single largest winning trade and the single largest losing trade.")
                if all_df.empty:
                    st.caption("No trades available.")
                else:
                    max_row = all_df.loc[all_df['Profit'].idxmax()]
                    min_row = all_df.loc[all_df['Profit'].idxmin()]

                    def trade_summary(row):
                        hold = row.get('Hold_Time', None)
                        return (f"{row.get('Symbol', 'N/A')} {row.get('Type', 'N/A')}, volume {row.get('Volume', 'N/A')}, "
                                f"held {fmt_td(hold) if hold is not None else 'N/A'}<br>"
                                f"Open: <b>{row.get('Open Time', 'N/A')}</b><br>Close: <b>{row.get('Close Time', 'N/A')}</b>")

                    c1, c2 = st.columns(2)
                    with c1:
                        status_banner(f"Largest win: ${float(max_row['Profit']):,.2f}", trade_summary(max_row), GOOD)
                    with c2:
                        status_banner(f"Largest loss: ${float(min_row['Profit']):,.2f}", trade_summary(min_row), BAD)

                    ext = pd.DataFrame([max_row, min_row])
                    ext.insert(0, 'Flag', ['Max Win', 'Max Loss'])
                    cols = ['Flag'] + [c for c in ['Symbol', 'Type', 'Open Time', 'Close Time', 'Volume', 'Profit', 'Hold_Time'] if c in ext.columns]
                    st.dataframe(ext[cols], use_container_width=True)
                    download_pair(ext[cols], f"max_win_loss_{account_no}", "flag_extremes", "MaxWinLoss", "Download")

        st.markdown("---")
        st.subheader("Client report")
        render_copy_button(_build_client_report(account_no, result, scalping_choice), f"single_{account_no}")

else:
    st.info("Upload one or more Excel reports to begin.")

st.markdown("---")

# ── IP lookup ──
st.header("IP lookup")
ip_input = st.text_area("IP addresses", placeholder="8.8.8.8, 1.1.1.1 or one per line", height=90, key="ip_input_field",
                        help="Separate addresses with commas or new lines.")
if st.button("Look up") and ip_input.strip():
    ip_list = [ip.strip() for ip in ip_input.replace("\n", ",").split(",") if ip.strip()]
    with st.spinner(f"Looking up {len(ip_list)} address(es)..."):
        for ip in ip_list:
            add_ip_to_history(ip, get_ip_details(ip))

if st.session_state.ip_history:
    st.markdown("### Recent lookups")
    entries = st.session_state.ip_history
    per_row = 3
    for start in range(0, len(entries), per_row):
        cols = st.columns(per_row)
        for col, entry in zip(cols, entries[start:start + per_row]):
            d = entry["details"]
            with col, st.container(border=True):
                if "error" in d:
                    st.markdown(f"**{entry['ip']}**")
                    st.error(d["error"])
                    continue
                st.markdown(f"**{entry['ip']}**")
                st.caption(f"{d.get('city', 'N/A')}, {d.get('region', 'N/A')}, {d.get('country', 'N/A')}")
                st.markdown(
                    f"ISP: {d.get('org', 'N/A')}  \n"
                    f"Timezone: {d.get('timezone', 'N/A')}  \n"
                    f"Coordinates: {d.get('loc', 'N/A')}"
                )
                st.caption(entry["timestamp"])
    if st.button("Clear history"):
        st.session_state.ip_history = []
        st.rerun()
