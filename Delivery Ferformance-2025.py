import streamlit as st
import pandas as pd
import numpy as np
import os 
import plotly.graph_objs as go
import re # Added for robust Part No. filtering

# --- CONFIG ---
# EXCEL_FILE = "Production-2025.xlsx" # Use local file name
EXCEL_FILE = "https://docs.google.com/spreadsheets/d/1C4An9Djq9j678VG8Hsz7mQlK57r8PTDU/export?format=xlsx"
# --------------

# -------------------------
# Helper: formatting
# -------------------------
def fmt_int(v):
    try:
        # Use .0f to ensure integers are displayed without trailing zeros
        return f"{int(v):,}"
    except Exception:
        return "0"

def fmt_pct(v):
    try:
        return f"{float(v):.2f}%"
    except Exception:
        return "0.00%"

# -------------------------
# Block 1: Cached Excel loader (FINAL, EXPLICIT RENAME FIX)
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
                nrows=max_rows
            )
            df.columns = df.columns.map(lambda x: str(x).strip())
            
            wk_num = int(sheet)
            new_edi_col_name = f"EDI_WK{wk_num:02d}"

            if "Part no." in df.columns:
                df['Part no.'] = df['Part no.'].astype(str).str.strip().str.upper()
                
                # --- CRITICAL RENAME LOGIC ---
                # Check for common names first
                expected_col_candidates = [f"WK{wk_num:02d}", f"W{wk_num}", str(wk_num)]
                found_col_name = next((c for c in expected_col_candidates if c in df.columns), None)

                if found_col_name:
                    # Rename the EDI column to a UNIQUE name to prevent collision
                    df.rename(columns={found_col_name: new_edi_col_name}, inplace=True)
                # --- END RENAME LOGIC ---
                
            dfs[sheet] = df
        except Exception as e:
            st.warning(f"Could not load sheet '{sheet}' or reach file: {e}")
    return dfs
# -------------------------
# Normalize columns
# -------------------------
def normalize_columns(df: pd.DataFrame):
    df = df.copy()
    df.columns = df.columns.map(lambda x: str(x).strip())
    return df

# -------------------------
# Sidebar Inputs (Week Range)
# -------------------------
st.sidebar.header("Select Week Range")
StartWeek = st.sidebar.selectbox('Start Week', [str(i) for i in range(1, 53)], index=3) # Changed index for common testing range
EndWeek = st.sidebar.selectbox('End Week', [str(i) for i in range(1, 53)], index=4)

StartWeek_int = int(StartWeek)
EndWeek_int = int(EndWeek)

if StartWeek_int > EndWeek_int:
    st.error("Start Week cannot be greater than End Week.")
    st.stop()

# -------------------------
# Block 2: Data Loading and Initial Merge (STREAMLINED & ROBUST)
# -------------------------
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

# 1. Assign Source_Week and Concatenate
DataMerges_list = []
week_map = {sheet: int(sheet) for sheet in dataframes.keys()}
for wk, df in dataframes.items():
    df_copy = df.copy() 
    df_copy["Source_Week"] = week_map[wk]
    DataMerges_list.append(df_copy)

DataMerges = pd.concat(DataMerges_list, ignore_index=True)

# 2. FINAL ROBUST DEDUPLICATION AND CLEANUP
if "Part no." in DataMerges.columns:
    
    # Drop duplicates by Part no. + Source_Week (This is guaranteed to work now 
    # because Part no. was cleaned inside load_dataframes).
    DataMerges.drop_duplicates(subset=["Part no.", "Source_Week"], keep="last", inplace=True)
    
    # 3. Final Filtering (Remaining steps are fine)
    DataMerges.dropna(subset=['Part no.'], inplace=True)

    # Use Regex to discard non-part-like identifiers (Keeps the 873 fix)
    DataMerges = DataMerges[
        DataMerges['Part no.'].str.contains(r'[A-Z0-9]', na=False) 
    ].copy()

    # Filter out known manual exclusions (Keeps your exclusion list)
    exclued = ['KOSHIN','HOME EXPERT','ELECTROLUX','TBKK','CAVAGNA','Thai GMB','1050B375','Z0011949A','Z0011951A','Z0020588A']
    DataMerges = DataMerges[~DataMerges['Part no.'].isin(exclued)].copy()
    
    # Filter out customer headers that may have slipped through (TBBK, ELECTROLUX)
    customer_headers = ['TBBK', 'ELECTROLUX', 'KOSHIN', 'HOME EXPERT', 'TOTAL', 'TOTAL PCS', 'SUMMARY'] 
    DataMerges = DataMerges[~DataMerges['Part no.'].isin(customer_headers)].copy()
    
else:
    st.error("Cannot process data: 'Part no.' column is missing after merging sheets.")
    st.stop()

# -------------------------
# DC Production (This section follows Block 2)
# -------------------------
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
# Stock setup (Standardizing ERO/Sales names)
# -------------------------
stock = normalize_columns(DataMerges.copy())
# Ensure 'Total' column is clean and numeric
if 'TOTAL' not in stock.columns and 'Total' in stock.columns:
    stock.rename(columns={'Total': 'TOTAL'}, inplace=True)
if 'TOTAL' in stock.columns:
    stock['TOTAL'] = pd.to_numeric(stock['TOTAL'], errors='coerce').fillna(0)

# Standardize ERO and Sales column names (Assumes 'ERO.6' and 'ACT.7' are correct)
stock.rename(columns={'ERO.6': 'ERO-Pcs', 'ACT.7': 'Sales-Pcs'}, inplace=True)
for col in ['ERO-Pcs', 'Sales-Pcs']:
    if col not in stock.columns:
        stock[col] = 0
    stock[col] = pd.to_numeric(stock[col], errors='coerce').fillna(0)


# 1. Identify ALL columns corresponding to the week range:
edi_cols_to_sum = []
all_cols = [str(c) for c in stock.columns]

for wk_num in week_numbers:
    # Look for the newly renamed, non-colliding column: EDI_WK04, EDI_WK05, etc.
    expected_col_name = f"EDI_WK{wk_num:02d}"
    
    if expected_col_name in all_cols:
        edi_cols_to_sum.append(expected_col_name)

# st.markdown(f"**Final EDI Columns Summed (RENAMED):** {edi_cols_to_sum}")
################################################################################
if not edi_cols_to_sum:
    st.error("EDI calculation failed: Could not find week forecast columns (e.g., EDI_WK04, EDI_WK05).")
    st.stop()
    
# 2. Convert identified EDI columns to numeric and sum across the row 
for col in edi_cols_to_sum:
    stock[col] = pd.to_numeric(stock[col], errors='coerce').fillna(0)
    
# CRITICAL EDI LINE: Calculate the total EDI for each ROW
stock['EDI-Total-Pcs'] = stock[edi_cols_to_sum].sum(axis=1)

# ... (Rest of the block is fine)

# ... (Rest of the block follows)


# 3. Aggregate EDI, ERO, and Sales by Part no.
SUMALL = stock.groupby('Part no.', as_index=False).agg({
    'EDI-Total-Pcs': 'sum',
    'ERO-Pcs': 'sum',
    'Sales-Pcs': 'sum'
})
SUMALL.set_index('Part no.', inplace=True)
SUMALL = SUMALL.fillna(0) 

# Recalculate GAP
SUMALL['GAP'] = SUMALL['ERO-Pcs'] - SUMALL['EDI-Total-Pcs']
SUMALL = SUMALL[~(SUMALL[['EDI-Total-Pcs', 'ERO-Pcs', 'GAP']] == 0).all(axis=1)]
SUMALL = SUMALL.merge(DCProd, left_index=True, right_index=True, how='left')

# st.markdown(f"**EDI Columns Summed:** {edi_cols_to_sum}")
# -------------------------
# Aggregates
# -------------------------
sum_FC = SUMALL['EDI-Total-Pcs'].sum() # Correct total EDI for the range
sum_ERO = SUMALL['ERO-Pcs'].sum()
sum_SALES = SUMALL['Sales-Pcs'].sum() 
# ... (rest of the aggregates remains the same)
SUM_MEET_ERO = SUMALL['ERO-Pcs'].where(SUMALL['GAP'] == 0).sum()
SUM_NoFC = SUMALL['ERO-Pcs'].where(SUMALL['EDI-Total-Pcs'] == 0).sum()
SUM_NO_ERO = SUMALL['EDI-Total-Pcs'].where(SUMALL['ERO-Pcs'] == 0).sum()
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
# Summary Table and Chart (Standard Code)
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

# ... (Chart code goes here, it remains unchanged from your original working code)
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
    if y_vals:
        fig.update_yaxes(range=[min(y_vals)*1.3, max(y_vals)*1.25])
    st.plotly_chart(fig, use_container_width=True)
except Exception as e:
    st.warning(f"Could not build plot: {e}")
    
# ... (Non-Delivery Order Report section follows, it also remains unchanged)
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

    ########################################################
