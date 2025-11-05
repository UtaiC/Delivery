import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import plotly.graph_objs as go
from pathlib import Path

# -------------------------
# UI: logo
# -------------------------
# try:
#     Logo = Image.open('SIM-LOGO-02.jpg')
#     st.image(Logo, width=720)
# except Exception:
#     st.warning("Logo not found or couldn't be opened.")

# -------------------------
# Helper: formatting
# -------------------------
def fmt_int(v):
    try:
        return f"{int(v):,}"
    except Exception:
        return "0"

def fmt_pct(v):
    try:
        return f"{float(v):.2f}%"
    except Exception:
        return "0.00%"

# -------------------------
# Cached Excel loader
# -------------------------
# -------------------------
# Cached Excel loader (limited rows)
# -------------------------
@st.cache_data
def load_dataframes(file_path: str, sheet_names: list, max_rows: int = 95):
    dfs = {}
    for sheet in sheet_names:
        try:
            df = pd.read_excel(
                file_path,
                header=7,
                engine='openpyxl',
                sheet_name=sheet,
                nrows=max_rows  # ✅ Limit to first 100 rows
            )
            dfs[sheet] = df
        except Exception as e:
            st.warning(f"Could not load sheet '{sheet}': {e}")
    return dfs

# -------------------------
# Normalize columns
# -------------------------
def normalize_columns(df: pd.DataFrame):
    df = df.copy()
    df.columns = df.columns.map(lambda x: str(x).strip())
    return df

# -------------------------
# Week finder + conversion
# -------------------------
def safe_find_and_numeric_week(df: pd.DataFrame, wk_key: str):
    if not isinstance(wk_key, str):
        wk_key = str(wk_key)
    wk_key = wk_key.strip().upper()
    digits = ''.join(ch for ch in wk_key if ch.isdigit())
    if digits == '':
        return None
    try:
        num = int(digits)
    except Exception:
        return None

    candidates = [
        f"WK{num:02d}", f"W{num}", f"W{num:02d}", f"WK{num}"
    ]
    cols = [str(c) for c in df.columns]
    found = next((c for c in candidates if c in cols), None)
    if found:
        df[found] = pd.to_numeric(df[found], errors='coerce')
        return found
    else:
        fuzzy = next((c for c in cols if digits in c and c.upper().startswith('W')), None)
        if fuzzy:
            df[fuzzy] = pd.to_numeric(df[fuzzy], errors='coerce')
            return fuzzy
    return None

# -------------------------
# Sidebar Inputs (Week Range)
# -------------------------
st.sidebar.header("Select Week Range")
StartWeek = st.sidebar.selectbox('Start Week', [str(i) for i in range(1, 53)], index=0)
EndWeek = st.sidebar.selectbox('End Week', [str(i) for i in range(1, 53)], index=4)

StartWeek_int = int(StartWeek)
EndWeek_int = int(EndWeek)

if StartWeek_int > EndWeek_int:
    st.error("Start Week cannot be greater than End Week.")
    st.stop()

# -------------------------
# Load Data from multiple weeks (with duplicate protection)
# -------------------------
# EXCEL_FILE = "Production-2025-WK45.xlsx"
EXCEL_FILE = "https://docs.google.com/spreadsheets/d/1C4An9Djq9j678VG8Hsz7mQlK57r8PTDU/export?format=xlsx"

# ensure unique week numbers even if Start == End
week_numbers = sorted(set(range(StartWeek_int, EndWeek_int + 1)))
sheets_to_load = [str(i) for i in week_numbers]

dataframes = load_dataframes(EXCEL_FILE, sheets_to_load)

if not dataframes:
    st.error(f"No data loaded for week range {StartWeek}–{EndWeek}. Check file: {EXCEL_FILE}")
    st.stop()

if StartWeek_int == EndWeek_int:
    st.subheader(f"EDI & EOR Report Week# {StartWeek}")
else:
    st.subheader(f"EDI & EOR Report Week# {StartWeek} → {EndWeek}")

# Merge all selected weeks
DataMerges = pd.concat(dataframes.values(), ignore_index=True)

# Optional: Add column to identify source week
week_map = {sheet: int(sheet) for sheet in dataframes.keys()}
for wk, df in dataframes.items():
    df["Source_Week"] = week_map[wk]

DataMerges = pd.concat(dataframes.values(), ignore_index=True)

# Drop duplicates if any rows have same key values (Part no., Week, etc.)
if "Part no." in DataMerges.columns:
    DataMerges.drop_duplicates(subset=["Part no.", "Source_Week"], keep="last", inplace=True)
else:
    DataMerges.drop_duplicates(inplace=True)


# -------------------------
# Exclude list
# -------------------------
exclued = ['Part no.','KOSHIN','HOME EXPERT','ELECTROLUX','TBKK','CAVAGNA','Thai GMB','1050B375','1050B375','Z0011949A','Z0011951A','Z0020588A']
DataMerges = DataMerges[~DataMerges['Part no.'].isin(exclued)].copy()
# -------------------------
# DC Production
# -------------------------
DCProd = DataMerges.copy()
for col in ['Total', 'Beginning Balance']:
    if col not in DCProd.columns:
        DCProd[col] = np.nan
DCProd['Total'] = pd.to_numeric(DCProd['Total'], errors='coerce')
DCProd['Beginning Balance'] = pd.to_numeric(DCProd['Beginning Balance'], errors='coerce')
DCProd['DC-Pcs'] = DCProd['Total'] - DCProd['Beginning Balance']
DCProd = DCProd[['Part no.', 'DC-Pcs']].groupby('Part no.', as_index=False).sum().set_index('Part no.')

# -------------------------
# Stock setup
# -------------------------
stock = normalize_columns(DataMerges.copy())
if 'TOTAL' not in stock.columns and 'Total' in stock.columns:
    stock.rename(columns={'Total': 'TOTAL'}, inplace=True)
if 'TOTAL' in stock.columns:
    stock['TOTAL'] = pd.to_numeric(stock['TOTAL'], errors='coerce').fillna(0)

# Use end week as active for ratio calc
active_week_col = safe_find_and_numeric_week(stock, f"WK{EndWeek_int:02d}")
if active_week_col is None:
    st.warning(f"Could not find column for WK{EndWeek_int}.")
    st.stop()

stock.rename(columns={'ERO.6': 'ERO-Pcs', 'ACT.7': 'Sales-Pcs'}, inplace=True)
for col in ['ERO-Pcs', 'Sales-Pcs']:
    if col not in stock.columns:
        stock[col] = 0
    stock[col] = pd.to_numeric(stock[col], errors='coerce').fillna(0)

stock[active_week_col] = pd.to_numeric(stock[active_week_col], errors='coerce').fillna(0)
stock['St-BL'] = stock['TOTAL'] - stock[active_week_col]
stock['ERO-FC'] = stock['ERO-Pcs'] - stock[active_week_col]
stock['ERO-%'] = np.where(stock[active_week_col] != 0, (stock['ERO-FC'] / stock[active_week_col]) * 100, 0)

# -------------------------
# Summary merge
# -------------------------
SUMALL = stock[['Part no.', active_week_col, 'ERO-Pcs']].copy()
SUMALL.rename(columns={active_week_col: 'EDI-' + active_week_col}, inplace=True)
SUMALL['GAP'] = SUMALL['ERO-Pcs'] - SUMALL['EDI-' + active_week_col]
SUMALL.set_index('Part no.', inplace=True)
SUMALL = SUMALL[~(SUMALL[['EDI-' + active_week_col, 'ERO-Pcs', 'GAP']] == 0).all(axis=1)]
SUMALL = SUMALL.merge(DCProd, left_index=True, right_index=True, how='left')

# -------------------------
# Aggregates
# -------------------------
sum_FC = SUMALL['EDI-' + active_week_col].sum()
sum_ERO = SUMALL['ERO-Pcs'].sum()
sum_SALES = stock['Sales-Pcs'].sum()
SUM_MEET_ERO = SUMALL['ERO-Pcs'].where(SUMALL['GAP'] == 0).sum()
SUM_NoFC = SUMALL['ERO-Pcs'].where(SUMALL['EDI-' + active_week_col] == 0).sum()
SUM_NO_ERO = SUMALL['EDI-' + active_week_col].where(SUMALL['ERO-Pcs'] == 0).sum()
SUM_LESS_ERO = SUMALL['ERO-Pcs'].where(SUMALL['GAP'] < 0).sum()
SUM_OVER_ERO = SUMALL['ERO-Pcs'].where(SUMALL['GAP'] > 0).sum()

# PCT calc
def pct(a, b): return (a / b * 100) if b != 0 else 0
EROPCT = pct(sum_ERO, sum_FC)
SalesPCT = pct(sum_SALES, sum_ERO)
MeetPCT = pct(SUM_MEET_ERO, sum_FC)
NoFC_PCT = pct(SUM_NoFC, sum_FC)
NOPCT = pct(SUM_NO_ERO, sum_FC)
LessPCT = pct(SUM_LESS_ERO, sum_FC)
OverPCT = pct(SUM_OVER_ERO, sum_FC)

# -------------------------
# Summary Table
# -------------------------
summary_items = [
    ('Forecasted (EDI)', sum_FC, '100%'),
    ('Ordered (ERO)', sum_ERO, EROPCT),
    ('Delivery (Sales)', sum_SALES, SalesPCT),
    ('EDI/ERO-Orders Met', SUM_MEET_ERO, MeetPCT),
    ('No EDI but EOR', SUM_NoFC, NoFC_PCT),
    ('No EOR but EDI', -SUM_NO_ERO, -NOPCT),
    ('Under-ERO (OVER EDI)', -SUM_LESS_ERO, -LessPCT),
    ('Over-ERO (Under EDI)', SUM_OVER_ERO, OverPCT),
]

df_summary = pd.DataFrame({
    'Row No': range(1, len(summary_items) + 1),
    'Item': [r[0] for r in summary_items],
    'Volumes (Pcs)': [r[1] for r in summary_items],
    'PCT-%': [r[2] for r in summary_items]
}).set_index('Row No')

df_display = df_summary.copy()
df_display['Volumes (Pcs)'] = df_display['Volumes (Pcs)'].apply(fmt_int)
df_display['PCT-%'] = df_display['PCT-%'].apply(lambda v: v if isinstance(v, str) else fmt_pct(v))

st.table(df_display)

# -------------------------
# Chart
# -------------------------
try:
    y_vals = [float(v) for v in df_summary['Volumes (Pcs)']]
    labels = [f"{fmt_int(v)}\n(PCT: {fmt_pct(p) if not isinstance(p, str) else p})" for v, p in zip(df_summary['Volumes (Pcs)'], df_summary['PCT-%'])]
    colors = ['#8CEC12','#ECD812','#12DBEC','#EC8C12','#EC8C12','#EC6E35','#EC4A12','#EC3312'][:len(y_vals)]

    fig = go.Figure(go.Bar(
        x=df_summary['Item'].tolist(),
        y=y_vals,
        text=labels,
        texttemplate="%{text}",
        textposition="outside",
        marker_color=colors,
    ))
    fig.update_traces(
        textposition=["outside" if y >= 0 else "inside" for y in y_vals],
        insidetextanchor="end"
    )
    fig.update_layout(
        title=f"Chart Summary EDI vs ERO (Pcs/PCT-%) WK{StartWeek}→WK{EndWeek}",
        xaxis_title="",
        yaxis_title="Value",
        height=720,
        margin=dict(t=140, b=150, l=60, r=60)
    )
    fig.update_yaxes(range=[min(y_vals)*1.3, max(y_vals)*1.25])
    st.plotly_chart(fig, use_container_width=True)
except Exception as e:
    st.warning(f"Could not build plot: {e}")
    
# Ensure numeric conversion first
for col in ['ERO-Pcs', 'Sales-Pcs']:
    if col in stock.columns:
        stock[col] = pd.to_numeric(stock[col], errors='coerce').fillna(0)

# If Sales = 0 but ERO > 0 → mark as repleted
stock['Non Delivery Order Pcs'] = np.where(
    (stock['Sales-Pcs'] == 0) & (stock['ERO-Pcs'] > 0),
    stock['ERO-Pcs'],
    0
)

# Create report for only >0
Repleted_Report = stock.loc[
    stock['Non Delivery Order Pcs'] > 0,
    ['Part no.', 'ERO-Pcs', 'Non Delivery Order Pcs']
].copy()

# --- Force all numeric columns to numeric ---
for col in ['ERO-Pcs', 'Non Delivery Order Pcs']:
    Repleted_Report[col] = pd.to_numeric(Repleted_Report[col], errors='coerce').fillna(0)

if not Repleted_Report.empty:
    st.subheader(f"🔁 Non Delivery Active Order Report (Week {StartWeek} → {EndWeek})")

    # --- Format numbers safely ---
    def safe_fmt(x):
        try:
            return f"{float(x):,.0f}"
        except Exception:
            return str(x)

    st.dataframe(Repleted_Report.applymap(safe_fmt))

    # --- Add total line ---
    SumReorder = float(Repleted_Report['Non Delivery Order Pcs'].sum())
   
    TotalNonSales = float( sum_SALES-sum_ERO)
    st.markdown(f"**Total Non Delivery Order Pcs:** {-SumReorder:,.0f}")
    st.markdown(f"**Total Missing Delivery Order Pcs:** {TotalNonSales:,.0f}")

else:
    st.info("No Repleted Orders detected in selected week range.")
#################


