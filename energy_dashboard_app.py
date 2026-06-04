import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import openpyxl
from datetime import datetime
import io
import os

# ── PDF imports ──────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from reportlab.platypus import Image as RLImage
import tempfile

# ═══════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Energy Intelligence · Sapphire NeelaBlue",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ═══════════════════════════════════════════════════════════════════════════
# THEME / CSS
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

/* ── Background ── */
.stApp {
    background: #0a0e1a;
    color: #e8eaf0;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #0d1220 !important;
    border-right: 1px solid #1e2a42;
}
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] label {
    color: #8892a4 !important;
    font-size: 0.78rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: #e8eaf0 !important;
    font-family: 'Syne', sans-serif;
}

/* ── Metric cards ── */
[data-testid="stMetric"] {
    background: #111827;
    border: 1px solid #1e2a42;
    border-radius: 12px;
    padding: 18px 20px !important;
    position: relative;
    overflow: hidden;
}
[data-testid="stMetric"]::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, #00d4ff, #7c3aed);
}
[data-testid="stMetricLabel"] {
    color: #8892a4 !important;
    font-size: 0.72rem !important;
    letter-spacing: 0.1em;
    text-transform: uppercase;
}
[data-testid="stMetricValue"] {
    color: #e8eaf0 !important;
    font-family: 'Syne', sans-serif !important;
    font-size: 1.6rem !important;
    font-weight: 700 !important;
}
[data-testid="stMetricDelta"] {
    font-size: 0.75rem !important;
}

/* ── Section headers ── */
.section-header {
    font-family: 'Syne', sans-serif;
    font-size: 1.05rem;
    font-weight: 700;
    color: #e8eaf0;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    margin: 0 0 4px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid #1e2a42;
}

/* ── Page title ── */
.page-title {
    font-family: 'Syne', sans-serif;
    font-size: 2rem;
    font-weight: 800;
    color: #e8eaf0;
    letter-spacing: -0.02em;
    line-height: 1.1;
}
.page-subtitle {
    font-family: 'DM Sans', sans-serif;
    font-size: 0.85rem;
    color: #4a5568;
    margin-top: 2px;
    letter-spacing: 0.06em;
    text-transform: uppercase;
}

/* ── Accent badge ── */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
.badge-blue  { background: #0d2240; color: #00d4ff; border: 1px solid #0a3a5c; }
.badge-green { background: #0d2a1a; color: #00e676; border: 1px solid #0a3a20; }
.badge-amber { background: #2a1a05; color: #ffab40; border: 1px solid #3a2a0a; }

/* ── Plotly chart container ── */
.chart-card {
    background: #111827;
    border: 1px solid #1e2a42;
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 16px;
}

/* ── Selectbox / slider ── */
[data-testid="stSelectbox"] > div,
[data-testid="stMultiSelect"] > div {
    background: #111827 !important;
    border-color: #1e2a42 !important;
    border-radius: 8px !important;
    color: #e8eaf0 !important;
}
.stSlider > div { color: #e8eaf0; }

/* ── Number input ── */
[data-testid="stNumberInput"] input {
    background: #111827 !important;
    border-color: #1e2a42 !important;
    color: #e8eaf0 !important;
    border-radius: 8px !important;
}

/* ── Buttons ── */
.stButton > button {
    background: linear-gradient(135deg, #0d2a4a, #1a1040) !important;
    border: 1px solid #1e3a5a !important;
    color: #00d4ff !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    border-radius: 8px !important;
    padding: 10px 20px !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    border-color: #00d4ff !important;
    box-shadow: 0 0 16px rgba(0,212,255,0.15) !important;
}

/* ── Download button ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #003d2a, #001a3a) !important;
    border: 1px solid #00e676 !important;
    color: #00e676 !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    border-radius: 8px !important;
    width: 100% !important;
}
[data-testid="stDownloadButton"] > button:hover {
    box-shadow: 0 0 16px rgba(0,230,118,0.2) !important;
}

/* ── Divider ── */
hr { border-color: #1e2a42 !important; }

/* ── Tabs ── */
[data-testid="stTabs"] [data-testid="stTab"] {
    font-family: 'Syne', sans-serif !important;
    font-weight: 600 !important;
    color: #4a5568 !important;
    letter-spacing: 0.05em;
    font-size: 0.8rem;
    text-transform: uppercase;
}
[data-testid="stTabs"] [aria-selected="true"] {
    color: #00d4ff !important;
    border-bottom-color: #00d4ff !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid #1e2a42;
    border-radius: 10px;
    overflow: hidden;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 6px; }
::-webkit-scrollbar-track { background: #0a0e1a; }
::-webkit-scrollbar-thumb { background: #1e2a42; border-radius: 3px; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# PLOTLY TEMPLATE
# ═══════════════════════════════════════════════════════════════════════════
PLOT_LAYOUT = dict(
    paper_bgcolor='#111827',
    plot_bgcolor='#111827',
    font=dict(family='DM Sans', color='#8892a4', size=11),
    xaxis=dict(gridcolor='#1e2a42', linecolor='#1e2a42', tickfont=dict(color='#8892a4')),
    yaxis=dict(gridcolor='#1e2a42', linecolor='#1e2a42', tickfont=dict(color='#8892a4')),
    legend=dict(bgcolor='#0d1220', bordercolor='#1e2a42', borderwidth=1,
                font=dict(color='#8892a4', size=10)),
    margin=dict(l=50, r=20, t=40, b=50),
    hovermode='x unified',
)

LT_COLORS = {
    'Rebeaming LT':              '#00d4ff',
    'Sizing LT':                 '#7c3aed',
    'Weaving LT-1':              '#00e676',
    'Weaving LT-2':              '#ffab40',
    'Weaving LT-3':              '#ff6b6b',
    'FINISHING LT':              '#f06292',
    'Finishing LT':              '#f06292',
    'Utility  BOILER LT':        '#80cbc4',
    'Utility AIR COMPRESSOR LT': '#ce93d8',
    'Utility AIR COMPRESSOR LT-2': '#fff176',
}

# ═══════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ═══════════════════════════════════════════════════════════════════════════
@st.cache_data
def load_data(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    records = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        date_str = sheet_name.strip()
        try:
            date = datetime.strptime(date_str, "%d-%m-%Y")
        except:
            continue

        rows = list(ws.iter_rows(values_only=True))
        current_lt = None
        lt_main = None
        lt_solar = None

        for row in rows:
            # LT section header: col0 is non-empty string, col1 is None, not a header row
            if (row[0] is not None and row[1] is None
                    and isinstance(row[0], str)
                    and len(row[0].strip()) > 2
                    and row[0].strip() not in ('Sr .#',)):
                # Save previous LT before switching
                if current_lt is not None and lt_main is not None:
                    records.append({
                        'date': date, 'lt': current_lt,
                        'main_kwh': max(lt_main, 0),
                        'solar_kwh': max(lt_solar or 0, 0)
                    })
                current_lt = row[0].strip()
                lt_main = None
                lt_solar = None
                continue

            if current_lt is None:
                continue

            # Main Breaker: always calculate from col D - col E (idx 3, 4)
            if row[1] == 'Main  Breaker':
                try:
                    t, y = row[3], row[4]
                    lt_main = (t - y) if isinstance(t, (int,float)) and isinstance(y, (int,float)) else 0
                except:
                    lt_main = 0

            # Solar row
            if row[1] is not None and 'Solar' in str(row[1]):
                try:
                    t, y = row[3], row[4]
                    lt_solar = (t - y) if isinstance(t, (int,float)) and isinstance(y, (int,float)) else 0
                except:
                    lt_solar = 0

        # Save last LT of sheet
        if current_lt is not None and lt_main is not None:
            records.append({
                'date': date, 'lt': current_lt,
                'main_kwh': max(lt_main, 0),
                'solar_kwh': max(lt_solar or 0, 0)
            })

    df = pd.DataFrame(records)
    df['lt'] = df['lt'].str.strip()
    df.loc[df['lt'] == 'Finishing LT', 'lt'] = 'FINISHING LT'
    df['main_kwh'] = df['main_kwh'].clip(lower=0)
    df['solar_kwh'] = df['solar_kwh'].clip(lower=0)
    df = df.sort_values('date').reset_index(drop=True)
    return df

# ── File uploader ──
uploaded_file = st.sidebar.file_uploader(
    "📂 Upload Load Report (.xlsx)",
    type=["xlsx"],
    help="Upload your Daily Load Report Excel file"
)

if uploaded_file is None:
    st.markdown("""
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;
                min-height:60vh;text-align:center;">
        <div style="background:#111827;border:1px solid #1e2a42;border-radius:18px;
                    padding:56px 64px;max-width:520px;">
            <p style="font-size:2.8rem;margin:0 0 16px;">⚡</p>
            <p style="font-family:Syne,sans-serif;font-size:1.6rem;color:#e8eaf0;
                      font-weight:800;margin:0 0 10px;letter-spacing:-0.02em;">
                Energy Intelligence</p>
            <p style="color:#4a5568;font-size:0.8rem;letter-spacing:0.1em;
                      text-transform:uppercase;margin:0 0 24px;">
                Sapphire Fibres · NeelaBlue Denim Unit</p>
            <div style="background:#0d1220;border:1px dashed #1e3a5a;border-radius:10px;
                        padding:20px 24px;margin-bottom:0;">
                <p style="color:#00d4ff;font-size:0.85rem;margin:0 0 6px;font-weight:600;">
                    📂 Upload your Daily Load Report</p>
                <p style="color:#8892a4;font-size:0.78rem;margin:0;">
                    Use the uploader in the sidebar to get started.<br>
                    Supports any month's Excel report.
                </p>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

df = load_data(uploaded_file.read())
ALL_LTS = sorted(df['lt'].unique().tolist())

# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('<p class="page-title" style="font-size:1.3rem;">⚡ NeelaBlue</p>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">Energy Intelligence Dashboard</p>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("**TARIFF SETTINGS**")
    blended_rate = st.number_input("Blended Rate — All Sources (Rs/kWh)", value=22.38, step=0.01, format="%.2f",
                                    help="Blended cost across LESCO + Solar + Gas Engine")
    lesco_rate   = st.number_input("LESCO-Only Rate (Rs/kWh)", value=46.00, step=0.01, format="%.2f",
                                    help="Pure LESCO tariff — used to calculate savings vs grid-only")
    st.markdown("---")

    st.markdown("**FILTERS**")
    selected_lts = st.multiselect("Select LT Substations", ALL_LTS, default=ALL_LTS,
                                   help="Filter which LTs to display")

    date_range = st.slider("Date Range (Day of May)",
                           min_value=1, max_value=31, value=(1, 31))
    st.markdown("---")

    st.markdown("**PDF REPORT**")
    pdf_lts = st.multiselect("LTs to include in Report", ALL_LTS, default=ALL_LTS)
    pdf_title = st.text_input("Report Title", value="Energy Report – May 2026")

# ═══════════════════════════════════════════════════════════════════════════
# FILTER DATA
# ═══════════════════════════════════════════════════════════════════════════
start_day = datetime(2026, 5, date_range[0])
end_day   = datetime(2026, 5, date_range[1])

fdf = df[
    (df['lt'].isin(selected_lts)) &
    (df['date'] >= start_day) &
    (df['date'] <= end_day)
].copy()

fdf['actual_cost']   = fdf['main_kwh'] * blended_rate
fdf['lesco_cost']    = fdf['main_kwh'] * lesco_rate
fdf['solar_savings'] = fdf['solar_kwh'] * (lesco_rate - blended_rate)
fdf['solar_offset_pct'] = fdf.apply(
    lambda r: round(r['solar_kwh'] / r['main_kwh'] * 100, 1) if r['main_kwh'] > 0 else 0, axis=1)

# ═══════════════════════════════════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════════════════════════════════
col_title, col_badge = st.columns([3, 1])
with col_title:
    st.markdown('<p class="page-title">Energy Intelligence</p>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">Sapphire Fibres · NeelaBlue Denim Unit · May 2026</p>', unsafe_allow_html=True)
with col_badge:
    st.markdown(f"""
    <div style="text-align:right; padding-top:10px;">
        <span class="badge badge-blue">⚡ {len(selected_lts)} LTs Active</span><br><br>
        <span class="badge badge-green">☀ Solar Monitored</span>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# KPI CARDS
# ═══════════════════════════════════════════════════════════════════════════
total_kwh    = fdf['main_kwh'].sum()
total_solar  = fdf['solar_kwh'].sum()
total_cost   = fdf['actual_cost'].sum()
total_savings= fdf['solar_savings'].sum()
avg_offset   = (total_solar / total_kwh * 100) if total_kwh > 0 else 0
peak_day_grp = fdf.groupby('date')['main_kwh'].sum()
peak_day     = peak_day_grp.idxmax().strftime("%-d May") if not peak_day_grp.empty else "—"
peak_kwh     = peak_day_grp.max() if not peak_day_grp.empty else 0

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Grid Consumption", f"{total_kwh:,.0f} kWh",
          delta=f"{total_kwh/1000:.1f} MWh total")
k2.metric("Total Solar Generated", f"{total_solar:,.0f} kWh",
          delta=f"{avg_offset:.1f}% avg offset")
k3.metric("Actual Cost (Blended)", f"Rs {total_cost/1e6:.2f}M",
          delta=f"Blended: LESCO + Solar + Gas @ Rs {blended_rate}")
k4.metric("Solar Savings vs LESCO", f"Rs {total_savings/1e6:.2f}M",
          delta=f"Savings vs LESCO-only @ Rs {lesco_rate}")
k5.metric("Peak Consumption Day", peak_day,
          delta=f"{peak_kwh:,.0f} kWh")

st.markdown("<br>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📈  TRENDS", "☀  SOLAR ANALYSIS", "💰  COST IMPACT", "📋  DATA TABLE", "🤖  ANOMALY DETECTION", "📊  FORECAST", "📝  AI REPORT"])

# ─── TAB 1: TRENDS ───────────────────────────────────────────────────────
with tab1:
    st.markdown('<p class="section-header">Daily Load Trend — All LTs</p>', unsafe_allow_html=True)

    daily_lt = fdf.groupby(['date', 'lt'])['main_kwh'].sum().reset_index()

    fig1 = go.Figure()
    for lt in selected_lts:
        sub = daily_lt[daily_lt['lt'] == lt]
        if sub.empty: continue
        c = LT_COLORS.get(lt, '#aaaaaa')
        fig1.add_trace(go.Scatter(
            x=sub['date'], y=sub['main_kwh'],
            name=lt, mode='lines',
            line=dict(color=c, width=2),
            fill='tozeroy',
            fillcolor='rgba({},{},{},0.06)'.format(int(c[1:3],16),int(c[3:5],16),int(c[5:7],16)) if c.startswith('#') else c,
            hovertemplate=f"<b>{lt}</b><br>%{{x|%d %b}}: %{{y:,.0f}} kWh<extra></extra>"
        ))
    fig1.update_layout(**PLOT_LAYOUT, height=360,
                       title=dict(text="Grid Consumption by LT (kWh/day)", font=dict(color='#e8eaf0', size=13, family='Syne')))
    st.plotly_chart(fig1, use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<p class="section-header">LT-wise Monthly Total (kWh)</p>', unsafe_allow_html=True)
        lt_total = fdf.groupby('lt')['main_kwh'].sum().sort_values(ascending=True).reset_index()
        fig2 = go.Figure(go.Bar(
            x=lt_total['main_kwh'], y=lt_total['lt'],
            orientation='h',
            marker=dict(
                color=[LT_COLORS.get(lt, '#aaaaaa') for lt in lt_total['lt']],
                opacity=0.85,
                line=dict(width=0)
            ),
            hovertemplate="<b>%{y}</b><br>%{x:,.0f} kWh<extra></extra>"
        ))
        fig2.update_layout(**PLOT_LAYOUT, height=320,
                           xaxis_title="kWh", yaxis_title="",
                           title=dict(text="Total Consumption Ranking", font=dict(color='#e8eaf0', size=12, family='Syne')))
        st.plotly_chart(fig2, use_container_width=True)

    with col_b:
        st.markdown('<p class="section-header">Daily Plant Total vs Solar</p>', unsafe_allow_html=True)
        daily_total = fdf.groupby('date').agg(main_kwh=('main_kwh','sum'), solar_kwh=('solar_kwh','sum')).reset_index()
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(x=daily_total['date'], y=daily_total['main_kwh'],
                              name='Grid Consumption', marker_color='#1e3a5a',
                              hovertemplate="%{x|%d %b}<br>Grid: %{y:,.0f} kWh<extra></extra>"))
        fig3.add_trace(go.Bar(x=daily_total['date'], y=daily_total['solar_kwh'],
                              name='Solar Generated', marker_color='#00d4ff',
                              hovertemplate="%{x|%d %b}<br>Solar: %{y:,.0f} kWh<extra></extra>"))
        fig3.update_layout(**PLOT_LAYOUT, barmode='overlay', height=320,
                           title=dict(text="Plant-wide: Grid vs Solar", font=dict(color='#e8eaf0', size=12, family='Syne')))
        st.plotly_chart(fig3, use_container_width=True)

# ─── TAB 2: SOLAR ────────────────────────────────────────────────────────
with tab2:
    st.markdown('<p class="section-header">Solar Generation & Offset Analysis</p>', unsafe_allow_html=True)

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        lt_solar_sum = fdf.groupby('lt')[['main_kwh','solar_kwh']].sum().reset_index()
        lt_solar_sum['offset_pct'] = (lt_solar_sum['solar_kwh'] / lt_solar_sum['main_kwh'] * 100).clip(0, 200).round(1)
        fig_s1 = go.Figure()
        fig_s1.add_trace(go.Bar(x=lt_solar_sum['lt'], y=lt_solar_sum['main_kwh'],
                                name='Grid', marker_color='#1e3a5a'))
        fig_s1.add_trace(go.Bar(x=lt_solar_sum['lt'], y=lt_solar_sum['solar_kwh'],
                                name='Solar', marker_color='#00d4ff'))
        fig_s1.update_layout(**PLOT_LAYOUT, barmode='group', height=340,
                             title=dict(text="Solar vs Grid per LT (May Total)", font=dict(color='#e8eaf0', size=12, family='Syne')),
                             xaxis_tickangle=-30)
        st.plotly_chart(fig_s1, use_container_width=True)

    with col_s2:
        fig_s2 = go.Figure(go.Bar(
            x=lt_solar_sum['lt'],
            y=lt_solar_sum['offset_pct'],
            marker=dict(
                color=lt_solar_sum['offset_pct'],
                colorscale=[[0,'#1e2a42'],[0.5,'#0d4a6a'],[1,'#00d4ff']],
                showscale=False,
                line=dict(width=0)
            ),
            hovertemplate="<b>%{x}</b><br>Solar Offset: %{y:.1f}%<extra></extra>"
        ))
        fig_s2.update_layout(**PLOT_LAYOUT, height=340,
                             title=dict(text="Solar Offset % per LT", font=dict(color='#e8eaf0', size=12, family='Syne')),
                             xaxis_tickangle=-30,
                             yaxis_title="Offset %")
        st.plotly_chart(fig_s2, use_container_width=True)

    st.markdown('<p class="section-header">Daily Solar Offset % by LT</p>', unsafe_allow_html=True)
    daily_offset = fdf.groupby(['date','lt']).agg(
        main=('main_kwh','sum'), solar=('solar_kwh','sum')).reset_index()
    daily_offset['pct'] = (daily_offset['solar'] / daily_offset['main'] * 100).clip(0,200).fillna(0)
    fig_s3 = go.Figure()
    for lt in selected_lts:
        sub = daily_offset[daily_offset['lt'] == lt]
        if sub.empty: continue
        c = LT_COLORS.get(lt, '#aaaaaa')
        fig_s3.add_trace(go.Scatter(
            x=sub['date'], y=sub['pct'], name=lt, mode='lines+markers',
            line=dict(color=c, width=1.5),
            marker=dict(size=4, color=c),
            hovertemplate=f"<b>{lt}</b><br>%{{x|%d %b}}: %{{y:.1f}}%<extra></extra>"
        ))
    fig_s3.add_hline(y=100, line_dash='dash', line_color='#00e676',
                     annotation_text="100% offset", annotation_font_color='#00e676')
    fig_s3.update_layout(**PLOT_LAYOUT, height=340,
                         title=dict(text="Daily Solar Offset % (Solar kWh / Grid kWh)", font=dict(color='#e8eaf0', size=12, family='Syne')),
                         yaxis_title="Offset %")
    st.plotly_chart(fig_s3, use_container_width=True)

# ─── TAB 3: COST ─────────────────────────────────────────────────────────
with tab3:
    st.markdown('<p class="section-header">Cost Impact & Solar Savings</p>', unsafe_allow_html=True)

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        lt_cost = fdf.groupby('lt').agg(
            actual=('actual_cost','sum'),
            lesco=('lesco_cost','sum'),
            savings=('solar_savings','sum')
        ).reset_index().sort_values('actual', ascending=False)
        fig_c1 = go.Figure()
        fig_c1.add_trace(go.Bar(x=lt_cost['lt'], y=lt_cost['lesco']/1000,
                                name='Without Solar/Gas (LESCO-only)', marker_color='#3a1a1a'))
        fig_c1.add_trace(go.Bar(x=lt_cost['lt'], y=lt_cost['actual']/1000,
                                name='Actual (All Sources Blended)', marker_color='#00d4ff'))
        fig_c1.update_layout(**PLOT_LAYOUT, barmode='overlay', height=340,
                             title=dict(text="Actual vs Without-Solar Cost per LT (Rs '000)", font=dict(color='#e8eaf0', size=12, family='Syne')),
                             xaxis_tickangle=-30,
                             yaxis_title="Rs '000")
        st.plotly_chart(fig_c1, use_container_width=True)

    with col_c2:
        fig_c2 = go.Figure(go.Bar(
            x=lt_cost['lt'], y=lt_cost['savings']/1000,
            marker=dict(color='#00e676', opacity=0.8),
            hovertemplate="<b>%{x}</b><br>Savings: Rs %{y:,.1f}K<extra></extra>"
        ))
        fig_c2.update_layout(**PLOT_LAYOUT, height=340,
                             title=dict(text="Solar Savings per LT (Rs '000)", font=dict(color='#e8eaf0', size=12, family='Syne')),
                             xaxis_tickangle=-30,
                             yaxis_title="Rs '000")
        st.plotly_chart(fig_c2, use_container_width=True)

    # Daily cost trend
    daily_cost = fdf.groupby('date').agg(
        actual=('actual_cost','sum'),
        savings=('solar_savings','sum')
    ).reset_index()
    fig_c3 = make_subplots(specs=[[{"secondary_y": True}]])
    fig_c3.add_trace(go.Bar(x=daily_cost['date'], y=daily_cost['actual']/1000,
                            name='Daily Cost (Rs K)', marker_color='#1e3a5a'), secondary_y=False)
    fig_c3.add_trace(go.Scatter(x=daily_cost['date'], y=daily_cost['savings']/1000,
                                name='Solar Savings (Rs K)', mode='lines',
                                line=dict(color='#00e676', width=2)), secondary_y=True)
    fig_c3.update_layout(**PLOT_LAYOUT, height=300,
                         title=dict(text="Daily Cost & Solar Savings Trend", font=dict(color='#e8eaf0', size=12, family='Syne')))
    fig_c3.update_yaxes(title_text="Cost (Rs K)", secondary_y=False, gridcolor='#1e2a42', tickfont=dict(color='#8892a4'))
    fig_c3.update_yaxes(title_text="Savings (Rs K)", secondary_y=True, gridcolor='#1e2a42', tickfont=dict(color='#00e676'))
    st.plotly_chart(fig_c3, use_container_width=True)

# ─── TAB 4: DATA TABLE ───────────────────────────────────────────────────
with tab4:
    st.markdown('<p class="section-header">Raw Data — LT-wise Daily Summary</p>', unsafe_allow_html=True)
    display_df = fdf[['date','lt','main_kwh','solar_kwh','solar_offset_pct','actual_cost','solar_savings']].copy()
    display_df['date'] = display_df['date'].dt.strftime('%d-%b-%Y')
    display_df.columns = ['Date','LT','Grid kWh','Solar kWh','Solar Offset %','Actual Cost (Rs)','Solar Savings (Rs)']
    display_df['Actual Cost (Rs)'] = display_df['Actual Cost (Rs)'].round(0)
    display_df['Solar Savings (Rs)'] = display_df['Solar Savings (Rs)'].round(0)
    st.dataframe(display_df, use_container_width=True, height=500)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown("---")

# ═══════════════════════════════════════════════════════════════════════════
# AI HELPER
# ═══════════════════════════════════════════════════════════════════════════
import json, numpy as np

def call_claude(system_prompt, user_prompt, max_tokens=2000):
    import urllib.request, urllib.error
    payload = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": max_tokens,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_prompt}]
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={"Content-Type": "application/json", "anthropic-version": "2023-06-01"},
        method="POST"
    )
    try:
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read())
        return data["content"][0]["text"]
    except urllib.error.HTTPError as e:
        return f"API Error {e.code}: {e.read().decode()}"
    except Exception as e:
        return f"Error: {str(e)}"

# ═══════════════════════════════════════════════════════════════════════════
# AI TABS (5, 6, 7)
# ═══════════════════════════════════════════════════════════════════════════

# ─── TAB 5: ANOMALY DETECTION ────────────────────────────────────────────
with tab5:
    st.markdown('<p class="section-header">AI Anomaly Detection</p>', unsafe_allow_html=True)
    st.markdown("""
    <p style="color:#8892a4;font-size:0.82rem;margin-bottom:18px;">
    AI analyzes each LT's daily consumption, flags statistically unusual days,
    and explains what likely caused each anomaly in engineering terms.
    </p>""", unsafe_allow_html=True)

    a_col1, a_col2 = st.columns([2,1])
    with a_col1:
        anomaly_lts = st.multiselect("Select LTs to analyze", ALL_LTS, default=ALL_LTS[:3], key="anom_lts")
    with a_col2:
        sensitivity = st.selectbox("Sensitivity", ["Low (2σ)", "Medium (1.5σ)", "High (1σ)"], index=1)
        sigma_map = {"Low (2σ)": 2.0, "Medium (1.5σ)": 1.5, "High (1σ)": 1.0}
        sigma = sigma_map[sensitivity]

    if st.button("🤖 Run Anomaly Detection", key="run_anomaly"):
        if not anomaly_lts:
            st.warning("Select at least one LT.")
        else:
            # Statistical detection
            anomalies = []
            for lt in anomaly_lts:
                lt_data = fdf[fdf['lt'] == lt].sort_values('date')
                if lt_data.empty or len(lt_data) < 5:
                    continue
                mean_kwh = lt_data['main_kwh'].mean()
                std_kwh  = lt_data['main_kwh'].std()
                if std_kwh == 0:
                    continue
                for _, row in lt_data.iterrows():
                    z = (row['main_kwh'] - mean_kwh) / std_kwh
                    if abs(z) >= sigma:
                        anomalies.append({
                            'lt': lt,
                            'date': row['date'].strftime('%d %b %Y'),
                            'kwh': row['main_kwh'],
                            'mean_kwh': round(mean_kwh, 0),
                            'z_score': round(z, 2),
                            'direction': 'HIGH' if z > 0 else 'LOW'
                        })

            if not anomalies:
                st.info("No anomalies detected at this sensitivity level. Try increasing sensitivity.")
            else:
                # Show chart with anomaly markers
                fig_a = go.Figure()
                for lt in anomaly_lts:
                    lt_data = fdf[fdf['lt'] == lt].sort_values('date')
                    if lt_data.empty: continue
                    c = LT_COLORS.get(lt, '#aaaaaa')
                    fig_a.add_trace(go.Scatter(
                        x=lt_data['date'], y=lt_data['main_kwh'],
                        name=lt, mode='lines',
                        line=dict(color=c, width=1.5),
                        hovertemplate=f"<b>{lt}</b><br>%{{x|%d %b}}: %{{y:,.0f}} kWh<extra></extra>"
                    ))
                # Add anomaly markers
                anom_df_plot = pd.DataFrame(anomalies)
                for _, a in anom_df_plot.iterrows():
                    color = '#ff4444' if a['direction'] == 'HIGH' else '#ffab40'
                    fig_a.add_trace(go.Scatter(
                        x=[pd.to_datetime(a['date'], format='%d %b %Y')],
                        y=[a['kwh']],
                        mode='markers',
                        marker=dict(color=color, size=12, symbol='circle', line=dict(color='white', width=1.5)),
                        name=f"⚠ {a['lt']} {a['date']}",
                        showlegend=False,
                        hovertemplate=f"<b>ANOMALY — {a['lt']}</b><br>{a['date']}: {a['kwh']:,.0f} kWh<br>z={a['z_score']}<extra></extra>"
                    ))
                fig_a.update_layout(**PLOT_LAYOUT, height=340,
                                    title=dict(text="Consumption with Anomaly Flags (red=high, amber=low)",
                                               font=dict(color='#e8eaf0', size=12, family='Syne')))
                st.plotly_chart(fig_a, use_container_width=True)

                st.markdown(f'<p class="section-header">Found {len(anomalies)} Anomalies — AI Analysis</p>', unsafe_allow_html=True)

                # Build data summary for AI
                anom_summary = json.dumps(anomalies, indent=2)
                lt_stats = {}
                for lt in anomaly_lts:
                    lt_data = fdf[fdf['lt'] == lt]
                    if not lt_data.empty:
                        lt_stats[lt] = {
                            'mean_kwh': round(lt_data['main_kwh'].mean(), 0),
                            'min_kwh':  round(lt_data['main_kwh'].min(), 0),
                            'max_kwh':  round(lt_data['main_kwh'].max(), 0),
                            'total_days': len(lt_data)
                        }

                with st.spinner("AI analyzing anomalies..."):
                    ai_response = call_claude(
                        system_prompt="""You are a Senior Electrical Engineer at a textile manufacturing plant (NeelaBlue Denim Unit, Sapphire Fibres).
You analyze power consumption anomalies across LT substations.
Your job: for each anomaly, provide a concise engineering explanation of likely causes.
Format your response as a structured list. For each anomaly:
- **[Date] — [LT Name] — [HIGH/LOW] — [kWh] kWh (z={z_score})**
  Likely cause: [1-2 sentence engineering explanation based on the LT type and direction]
  Action: [brief recommended action]

Consider: Weaving LTs power looms (high base load), Finishing LT has sanforizing/washing machines, Rebeaming LT has winding/creel machines, Sizing LT has sizing machines and kitchen, Utility LTs have boilers/compressors/chillers.
Be specific, use engineering terminology. No generic responses.""",
                        user_prompt=f"""Anomalies detected (sigma threshold={sigma}):
{anom_summary}

LT baseline stats:
{json.dumps(lt_stats, indent=2)}

Analyze each anomaly and explain probable engineering causes.""",
                        max_tokens=2000
                    )

                st.markdown(f"""
                <div style="background:#0d1220;border:1px solid #1e2a42;border-radius:12px;padding:20px 24px;
                            font-family:'DM Sans',sans-serif;font-size:0.85rem;line-height:1.7;color:#c0c8d8;">
                {ai_response.replace(chr(10), '<br>').replace('**', '<b>').replace('**', '</b>')}
                </div>""", unsafe_allow_html=True)

                # Raw anomaly table
                st.markdown("<br>", unsafe_allow_html=True)
                anom_display = pd.DataFrame(anomalies)
                anom_display.columns = ['LT','Date','kWh','Avg kWh','Z-Score','Type']
                anom_display['Type'] = anom_display['Type'].apply(lambda x: '🔴 HIGH' if x=='HIGH' else '🟡 LOW')
                st.dataframe(anom_display, use_container_width=True)

# ─── TAB 6: FORECAST ─────────────────────────────────────────────────────
with tab6:
    st.markdown('<p class="section-header">AI Consumption Forecast</p>', unsafe_allow_html=True)
    st.markdown("""
    <p style="color:#8892a4;font-size:0.82rem;margin-bottom:18px;">
    Linear trend + 7-day rolling forecast per LT with confidence bands.
    AI interprets the forecast and flags any LTs showing concerning trends.
    </p>""", unsafe_allow_html=True)

    f_col1, f_col2 = st.columns([2,1])
    with f_col1:
        forecast_lts = st.multiselect("Select LTs to forecast", ALL_LTS, default=ALL_LTS[:4], key="fore_lts")
    with f_col2:
        forecast_days = st.slider("Forecast horizon (days)", 3, 14, 7)

    if st.button("📊 Generate Forecast", key="run_forecast"):
        if not forecast_lts:
            st.warning("Select at least one LT.")
        else:
            from numpy.polynomial import polynomial as P

            fig_f = go.Figure()
            forecast_results = {}

            for lt in forecast_lts:
                lt_data = fdf[fdf['lt'] == lt].sort_values('date').reset_index(drop=True)
                if lt_data.empty or len(lt_data) < 7: continue
                c = LT_COLORS.get(lt, '#aaaaaa')

                x = np.arange(len(lt_data))
                y = lt_data['main_kwh'].values

                # Linear regression
                coeffs = np.polyfit(x, y, 1)
                trend_line = np.polyval(coeffs, x)
                residuals = y - trend_line
                std_resid = residuals.std()

                # Forecast future points
                x_future = np.arange(len(lt_data), len(lt_data) + forecast_days)
                y_future = np.polyval(coeffs, x_future)
                y_upper  = y_future + 1.5 * std_resid
                y_lower  = np.maximum(y_future - 1.5 * std_resid, 0)

                last_date = lt_data['date'].max()
                future_dates = pd.date_range(last_date + pd.Timedelta(days=1), periods=forecast_days)

                # Actual line
                fig_f.add_trace(go.Scatter(
                    x=lt_data['date'], y=lt_data['main_kwh'],
                    name=lt, mode='lines',
                    line=dict(color=c, width=2),
                    hovertemplate=f"<b>{lt}</b><br>%{{x|%d %b}}: %{{y:,.0f}} kWh<extra></extra>"
                ))
                # Forecast line
                fig_f.add_trace(go.Scatter(
                    x=future_dates, y=y_future,
                    name=f"{lt} forecast", mode='lines',
                    line=dict(color=c, width=2, dash='dash'),
                    showlegend=False,
                    hovertemplate=f"<b>{lt} FORECAST</b><br>%{{x|%d %b}}: %{{y:,.0f}} kWh<extra></extra>"
                ))
                # Confidence band
                r_hex = c[1:3]; g_hex = c[3:5]; b_hex = c[5:7]
                rgba_fill = f"rgba({int(r_hex,16)},{int(g_hex,16)},{int(b_hex,16)},0.1)"
                fig_f.add_trace(go.Scatter(
                    x=list(future_dates) + list(future_dates[::-1]),
                    y=list(y_upper) + list(y_lower[::-1]),
                    fill='toself', fillcolor=rgba_fill,
                    line=dict(color='rgba(0,0,0,0)'),
                    showlegend=False, hoverinfo='skip'
                ))

                forecast_results[lt] = {
                    'trend_slope': round(float(coeffs[0]), 1),
                    'avg_kwh': round(float(y.mean()), 0),
                    'forecast_7d_avg': round(float(y_future.mean()), 0),
                    'forecast_next': round(float(y_future[0]), 0),
                    'direction': 'increasing' if coeffs[0] > 50 else 'decreasing' if coeffs[0] < -50 else 'stable'
                }

            # Add vertical divider line
            last_actual = fdf['date'].max()
            fig_f.add_vline(x=last_actual, line_dash="dot", line_color="#4a5568",
                            annotation_text="Forecast →", annotation_font_color="#4a5568")
            fig_f.update_layout(**PLOT_LAYOUT, height=380,
                                title=dict(text=f"Consumption Trend + {forecast_days}-Day Forecast (dashed)",
                                           font=dict(color='#e8eaf0', size=12, family='Syne')))
            st.plotly_chart(fig_f, use_container_width=True)

            # Forecast summary table
            fc_rows = []
            for lt, v in forecast_results.items():
                arrow = "↑" if v['direction'] == 'increasing' else "↓" if v['direction'] == 'decreasing' else "→"
                fc_rows.append([lt, f"{v['avg_kwh']:,.0f}", f"{v['forecast_next']:,.0f}",
                                 f"{v['forecast_7d_avg']:,.0f}", f"{v['trend_slope']:+.1f} kWh/day", arrow + " " + v['direction']])
            fc_df = pd.DataFrame(fc_rows, columns=['LT','May Avg kWh','Tomorrow','7-Day Avg','Trend Slope','Direction'])
            st.dataframe(fc_df, use_container_width=True)

            # AI interpretation
            st.markdown('<p class="section-header">AI Forecast Interpretation</p>', unsafe_allow_html=True)
            with st.spinner("AI analyzing forecast..."):
                ai_fc = call_claude(
                    system_prompt="""You are a Senior Electrical Engineer at NeelaBlue Denim Unit, Sapphire Fibres.
You interpret energy consumption forecasts for LT substations.
Write a concise professional commentary (4-6 paragraphs) covering:
1. Overall plant consumption trend heading into next week
2. Which LTs show concerning increasing trends and why (in context of textile operations)
3. Which LTs show healthy/stable patterns
4. Estimated cost impact of the forecasted consumption (use blended rate provided)
5. One actionable recommendation per concerning LT
Use engineering terminology. Be specific, not generic.""",
                    user_prompt=f"""Forecast results for {forecast_days}-day horizon:
{json.dumps(forecast_results, indent=2)}

Blended tariff: Rs {blended_rate}/kWh
LESCO tariff: Rs {lesco_rate}/kWh

Provide a professional engineering interpretation.""",
                    max_tokens=1800
                )

            st.markdown(f"""
            <div style="background:#0d1220;border:1px solid #1e2a42;border-radius:12px;padding:20px 24px;
                        font-family:'DM Sans',sans-serif;font-size:0.85rem;line-height:1.7;color:#c0c8d8;">
            {ai_fc.replace(chr(10), '<br>')}
            </div>""", unsafe_allow_html=True)

# ─── TAB 7: AI REPORT ────────────────────────────────────────────────────
with tab7:
    st.markdown('<p class="section-header">AI Executive Report Generator</p>', unsafe_allow_html=True)
    st.markdown("""
    <p style="color:#8892a4;font-size:0.82rem;margin-bottom:18px;">
    AI writes a full engineering commentary for the selected LTs and period,
    then embeds it into the PDF report alongside charts and data tables.
    </p>""", unsafe_allow_html=True)

    r_col1, r_col2 = st.columns(2)
    with r_col1:
        report_lts_ai = st.multiselect("LTs to include in report", ALL_LTS, default=ALL_LTS, key="rpt_lts")
        report_title  = st.text_input("Report Title", value="Energy Intelligence Report — May 2026", key="rpt_title")
    with r_col2:
        report_sections = st.multiselect(
            "Sections to include",
            ["Executive Summary", "LT-wise Analysis", "Solar Performance", "Cost Analysis", "Anomaly Review", "Recommendations"],
            default=["Executive Summary", "LT-wise Analysis", "Solar Performance", "Cost Analysis", "Recommendations"]
        )

    if st.button("📝 Generate AI Report", key="gen_report"):
        if not report_lts_ai:
            st.warning("Select at least one LT.")
        else:
            rpt_df = df[
                (df['lt'].isin(report_lts_ai)) &
                (df['date'] >= start_day) &
                (df['date'] <= end_day)
            ].copy()
            rpt_df['actual_cost']   = rpt_df['main_kwh'] * blended_rate
            rpt_df['lesco_cost']    = rpt_df['main_kwh'] * lesco_rate
            rpt_df['solar_savings'] = rpt_df['solar_kwh'] * (lesco_rate - blended_rate)

            # Build data summary for AI
            lt_summary = rpt_df.groupby('lt').agg(
                grid=('main_kwh','sum'), solar=('solar_kwh','sum'),
                cost=('actual_cost','sum'), savings=('solar_savings','sum')
            ).reset_index()
            lt_summary['offset_pct'] = (lt_summary['solar'] / lt_summary['grid'] * 100).round(1)
            lt_summary['peak_day'] = rpt_df.groupby('lt')['main_kwh'].max().values

            plant_total_kwh  = rpt_df['main_kwh'].sum()
            plant_total_solar= rpt_df['solar_kwh'].sum()
            plant_total_cost = rpt_df['actual_cost'].sum()
            plant_savings    = rpt_df['solar_savings'].sum()
            days_covered     = (end_day - start_day).days + 1

            data_for_ai = {
                "report_period": f"1–{date_range[1]} May 2026 ({days_covered} days)",
                "lts_covered": report_lts_ai,
                "tariffs": {"blended": blended_rate, "lesco_only": lesco_rate},
                "plant_totals": {
                    "grid_kwh": round(plant_total_kwh, 0),
                    "solar_kwh": round(plant_total_solar, 0),
                    "avg_solar_offset_pct": round(plant_total_solar/plant_total_kwh*100, 1) if plant_total_kwh > 0 else 0,
                    "total_cost_rs": round(plant_total_cost, 0),
                    "solar_savings_rs": round(plant_savings, 0),
                },
                "lt_breakdown": lt_summary.to_dict(orient='records'),
                "sections_requested": report_sections
            }

            with st.spinner("AI generating report narrative..."):
                ai_narrative = call_claude(
                    system_prompt="""You are a Senior Electrical Engineer writing an official monthly energy report for NeelaBlue Denim Unit, Sapphire Fibres Limited, Lahore.
Write a professional, detailed engineering report narrative covering the requested sections.
Format with clear section headings using markdown (## for sections, ### for subsections).
Be specific with numbers, percentages, and engineering observations.
Highlight concerns, achievements, and actionable recommendations.
Tone: professional engineering report, suitable for presentation to GM Engineering and management.
Use proper engineering terminology for textile/denim manufacturing equipment.""",
                    user_prompt=f"""Generate the report narrative based on this data:
{json.dumps(data_for_ai, indent=2, default=str)}

Sections requested: {report_sections}
Write a comprehensive engineering report.""",
                    max_tokens=2500
                )

            # Show narrative in app
            st.markdown('<p class="section-header">AI-Generated Narrative Preview</p>', unsafe_allow_html=True)
            st.markdown(f"""
            <div style="background:#0d1220;border:1px solid #1e2a42;border-radius:12px;padding:24px 28px;
                        font-family:'DM Sans',sans-serif;font-size:0.85rem;line-height:1.8;color:#c0c8d8;
                        max-height:500px;overflow-y:auto;">
            {ai_narrative.replace(chr(10), '<br>')}
            </div>""", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Generate PDF with AI narrative
            with st.spinner("Building PDF with AI narrative..."):
                pdf_buf = generate_pdf_ai(
                    rpt_df, report_lts_ai, report_title,
                    blended_rate, lesco_rate, ai_narrative, report_sections
                )

            fname = f"AI_Energy_Report_{datetime.now().strftime('%d%m%Y_%H%M%S')}.pdf"
            st.download_button(
                label="📥 Download AI Report PDF",
                data=pdf_buf,
                file_name=fname,
                mime="application/pdf"
            )

# ═══════════════════════════════════════════════════════════════════════════
# PDF GENERATION
# ═══════════════════════════════════════════════════════════════════════════
def make_mpl_chart(dates, series_dict, title, ylabel, colors_map, height_in=2.8):
    fig, ax = plt.subplots(figsize=(7.2, height_in), facecolor='#111827')
    ax.set_facecolor('#111827')
    for name, vals in series_dict.items():
        c = colors_map.get(name, '#aaaaaa')
        ax.plot(range(len(dates)), vals, color=c, linewidth=1.5, label=name)
    ax.set_xticks(range(0, len(dates), max(1, len(dates)//8)))
    ax.set_xticklabels([dates[i].strftime('%-d') for i in range(0, len(dates), max(1, len(dates)//8))],
                        color='#8892a4', fontsize=7)
    ax.tick_params(axis='y', colors='#8892a4', labelsize=7)
    ax.set_ylabel(ylabel, color='#8892a4', fontsize=7)
    ax.set_title(title, color='#e8eaf0', fontsize=9, fontweight='bold', pad=6)
    ax.spines[:].set_color('#1e2a42')
    ax.grid(axis='y', color='#1e2a42', linewidth=0.5)
    if len(series_dict) > 1:
        ax.legend(fontsize=6, framealpha=0.3, facecolor='#0d1220', edgecolor='#1e2a42',
                  labelcolor='#8892a4', loc='upper right', ncol=2)
    plt.tight_layout(pad=0.4)
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='#111827')
    plt.close()
    buf.seek(0)
    return buf


def _pdf_styles():
    title_style = ParagraphStyle('ptitle', fontName='Helvetica-Bold',
                                  fontSize=18, textColor=colors.HexColor('#e8eaf0'),
                                  spaceAfter=2, alignment=TA_LEFT)
    sub_style   = ParagraphStyle('psub', fontName='Helvetica',
                                  fontSize=8, textColor=colors.HexColor('#4a5568'),
                                  spaceAfter=12, alignment=TA_LEFT, leading=12)
    h2_style    = ParagraphStyle('ph2', fontName='Helvetica-Bold',
                                  fontSize=11, textColor=colors.HexColor('#00d4ff'),
                                  spaceBefore=14, spaceAfter=4)
    h3_style    = ParagraphStyle('ph3', fontName='Helvetica-Bold',
                                  fontSize=9.5, textColor=colors.HexColor('#ffab40'),
                                  spaceBefore=10, spaceAfter=3)
    body_style  = ParagraphStyle('pbody', fontName='Helvetica',
                                  fontSize=8.5, textColor=colors.HexColor('#c0c8d8'),
                                  leading=13)
    label_style = ParagraphStyle('plabel', fontName='Helvetica',
                                  fontSize=7.5, textColor=colors.HexColor('#8892a4'),
                                  alignment=TA_CENTER)
    val_style   = ParagraphStyle('pval', fontName='Helvetica-Bold',
                                  fontSize=13, textColor=colors.HexColor('#e8eaf0'),
                                  alignment=TA_CENTER)
    return title_style, sub_style, h2_style, h3_style, body_style, label_style, val_style


def _table_style_dark(has_total=False):
    style = [
        ('BACKGROUND',   (0,0), (-1,0),  colors.HexColor('#0d2240')),
        ('TEXTCOLOR',    (0,0), (-1,0),  colors.HexColor('#00d4ff')),
        ('FONTNAME',     (0,0), (-1,0),  'Helvetica-Bold'),
        ('FONTSIZE',     (0,0), (-1,-1), 7.5),
        ('BACKGROUND',   (0,1), (-1,-1), colors.HexColor('#111827')),
        ('TEXTCOLOR',    (0,1), (-1,-1), colors.HexColor('#c0c8d8')),
        ('ROWBACKGROUNDS',(0,1),(-1,-1), [colors.HexColor('#111827'), colors.HexColor('#0d1220')]),
        ('GRID',         (0,0), (-1,-1), 0.4, colors.HexColor('#1e2a42')),
        ('ALIGN',        (1,0), (-1,-1), 'RIGHT'),
        ('LEFTPADDING',  (0,0), (-1,-1), 7),
        ('RIGHTPADDING', (0,0), (-1,-1), 7),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
    ]
    if has_total:
        style += [
            ('FONTNAME',   (0,-1),(-1,-1), 'Helvetica-Bold'),
            ('BACKGROUND', (0,-1),(-1,-1), colors.HexColor('#0d2240')),
            ('TEXTCOLOR',  (0,-1),(-1,-1), colors.HexColor('#00e676')),
        ]
    return TableStyle(style)


def generate_pdf_ai(report_df, lt_list, title_str, blended, lesco, ai_narrative, sections):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=18*mm, rightMargin=18*mm,
                            topMargin=18*mm, bottomMargin=18*mm)
    W = A4[0] - 36*mm
    title_style, sub_style, h2_style, h3_style, body_style, label_style, val_style = _pdf_styles()
    story = []

    # ── Cover ──
    story.append(Paragraph(title_str, title_style))
    story.append(Paragraph(
        f"Sapphire Fibres Limited · NeelaBlue Denim Unit · Generated {datetime.now().strftime('%d %b %Y, %H:%M')}",
        sub_style))
    story.append(Paragraph(
        f"LTs Covered: {', '.join(lt_list)} · Blended Rs {blended}/kWh · LESCO Rs {lesco}/kWh",
        sub_style))
    story.append(HRFlowable(width=W, thickness=1, color=colors.HexColor('#1e2a42'), spaceAfter=10))

    # ── Plant Summary KPIs ──
    lt_sum = report_df.groupby('lt').agg(
        grid=('main_kwh','sum'), solar=('solar_kwh','sum'),
        cost=('actual_cost','sum'), savings=('solar_savings','sum')
    ).reset_index()
    lt_sum['offset_pct'] = (lt_sum['solar'] / lt_sum['grid'] * 100).clip(0,200).round(1)

    story.append(Paragraph("Plant Summary", h2_style))
    kpi_data = [['LT Substation', 'Grid (kWh)', 'Solar (kWh)', 'Offset %', 'Cost (Rs)', 'Savings (Rs)']]
    for _, row in lt_sum.iterrows():
        kpi_data.append([row['lt'], f"{row['grid']:,.0f}", f"{row['solar']:,.0f}",
                         f"{row['offset_pct']:.1f}%", f"{row['cost']:,.0f}", f"{row['savings']:,.0f}"])
    kpi_data.append(['TOTAL',
                     f"{lt_sum['grid'].sum():,.0f}", f"{lt_sum['solar'].sum():,.0f}",
                     f"{lt_sum['solar'].sum()/max(lt_sum['grid'].sum(),1)*100:.1f}%",
                     f"{lt_sum['cost'].sum():,.0f}", f"{lt_sum['savings'].sum():,.0f}"])
    t = Table(kpi_data, colWidths=[W*0.28,W*0.13,W*0.13,W*0.1,W*0.18,W*0.18])
    t.setStyle(_table_style_dark(has_total=True))
    story.append(t)
    story.append(Spacer(1, 12))

    # ── AI Narrative ──
    story.append(Paragraph("AI Engineering Analysis", h2_style))
    story.append(HRFlowable(width=W, thickness=0.5, color=colors.HexColor('#1e2a42'), spaceAfter=6))

    # Parse markdown headings into reportlab paragraphs
    for line in ai_narrative.split('\n'):
        line = line.strip()
        if not line:
            story.append(Spacer(1, 4))
        elif line.startswith('### '):
            story.append(Paragraph(line[4:], h3_style))
        elif line.startswith('## '):
            story.append(Paragraph(line[3:], h2_style))
        elif line.startswith('# '):
            story.append(Paragraph(line[2:], h2_style))
        elif line.startswith('- ') or line.startswith('* '):
            story.append(Paragraph(f"• {line[2:]}", body_style))
        elif line.startswith('**') and line.endswith('**'):
            story.append(Paragraph(f"<b>{line[2:-2]}</b>", body_style))
        else:
            # Handle inline bold
            line = line.replace('**', '<b>', 1)
            while '**' in line:
                line = line.replace('**', '</b>', 1)
            story.append(Paragraph(line, body_style))

    # ── Per-LT charts ──
    for lt in lt_list:
        story.append(PageBreak())
        story.append(Paragraph(f"LT Detail: {lt}", h2_style))
        story.append(HRFlowable(width=W, thickness=0.5, color=colors.HexColor('#1e2a42'), spaceAfter=8))

        lt_df = report_df[report_df['lt'] == lt].sort_values('date')
        if lt_df.empty:
            story.append(Paragraph("No data.", body_style))
            continue

        g = lt_df['main_kwh'].sum(); s = lt_df['solar_kwh'].sum()
        ac = lt_df['actual_cost'].sum(); sv = lt_df['solar_savings'].sum()
        op = s/g*100 if g > 0 else 0

        kpi_mini = [
            [Paragraph('Grid', label_style), Paragraph('Solar', label_style),
             Paragraph('Offset %', label_style), Paragraph('Cost', label_style),
             Paragraph('Savings', label_style)],
            [Paragraph(f"{g:,.0f} kWh", val_style), Paragraph(f"{s:,.0f} kWh", val_style),
             Paragraph(f"{op:.1f}%", val_style), Paragraph(f"Rs {ac:,.0f}", val_style),
             Paragraph(f"Rs {sv:,.0f}", val_style)],
        ]
        t_mini = Table(kpi_mini, colWidths=[W/5]*5)
        t_mini.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#111827')),
            ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor('#1e2a42')),
            ('INNERGRID',  (0,0), (-1,-1), 0.5, colors.HexColor('#1e2a42')),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING',(0,0),(-1,-1),6),
        ]))
        story.append(t_mini)
        story.append(Spacer(1, 8))

        lt_dates = lt_df['date'].tolist()
        fig_buf = make_mpl_chart(
            lt_dates,
            {'Grid kWh': lt_df['main_kwh'].tolist(), 'Solar kWh': lt_df['solar_kwh'].tolist()},
            f"{lt} — Daily Grid & Solar (kWh)", "kWh",
            {'Grid kWh': '#00d4ff', 'Solar kWh': '#00e676'}, height_in=2.4
        )
        story.append(RLImage(fig_buf, width=W, height=W*0.33))
        story.append(Spacer(1, 6))

        fig_buf2 = make_mpl_chart(
            lt_dates,
            {'Actual Cost (Rs)': lt_df['actual_cost'].tolist(), 'Solar Savings (Rs)': lt_df['solar_savings'].tolist()},
            f"{lt} — Daily Cost & Savings (Rs)", "Rs",
            {'Actual Cost (Rs)': '#ffab40', 'Solar Savings (Rs)': '#00e676'}, height_in=2.4
        )
        story.append(RLImage(fig_buf2, width=W, height=W*0.33))
        story.append(Spacer(1, 8))

        # Daily table
        tbl_data = [['Date','Grid kWh','Solar kWh','Offset %','Cost (Rs)','Savings (Rs)']]
        for _, r in lt_df.iterrows():
            op_r = r['solar_kwh']/r['main_kwh']*100 if r['main_kwh'] > 0 else 0
            tbl_data.append([r['date'].strftime('%d %b'), f"{r['main_kwh']:,.0f}",
                             f"{r['solar_kwh']:,.0f}", f"{op_r:.1f}%",
                             f"{r['actual_cost']:,.0f}", f"{r['solar_savings']:,.0f}"])
        t_d = Table(tbl_data, colWidths=[W*0.13,W*0.15,W*0.15,W*0.12,W*0.22,W*0.23])
        t_d.setStyle(_table_style_dark())
        story.append(t_d)

    doc.build(story)
    buf.seek(0)
    return buf


# Keep old generate_pdf for the standalone PDF section
def generate_pdf(report_df, lt_list, title_str, blended, lesco):
    return generate_pdf_ai(report_df, lt_list, title_str, blended, lesco, "", [])

# ═══════════════════════════════════════════════════════════════════════════
# STANDALONE PDF SECTION (below tabs)
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<p class="section-header">Quick PDF Export (No AI Narrative)</p>', unsafe_allow_html=True)

pdf_col1, pdf_col2 = st.columns([3, 1])
with pdf_col1:
    pdf_lts_quick = st.multiselect("LTs for quick export", ALL_LTS, default=ALL_LTS, key="quick_pdf_lts")
    st.markdown(f"""
    <div style="background:#111827; border:1px solid #1e2a42; border-radius:10px; padding:12px 18px; margin-top:8px;">
        <span style="color:#8892a4; font-size:0.75rem; letter-spacing:0.08em; text-transform:uppercase;">
        {len(pdf_lts_quick)} LT(s) · Days {date_range[0]}–{date_range[1]} May 2026 · 
        Blended Rs {blended_rate}/kWh · LESCO Rs {lesco_rate}/kWh</span>
    </div>""", unsafe_allow_html=True)

with pdf_col2:
    if st.button("⚡ Export PDF", key="quick_pdf_btn"):
        if not pdf_lts_quick:
            st.error("Select at least one LT.")
        else:
            with st.spinner("Building PDF..."):
                pdf_df = df[
                    (df['lt'].isin(pdf_lts_quick)) &
                    (df['date'] >= start_day) &
                    (df['date'] <= end_day)
                ].copy()
                pdf_df['actual_cost']   = pdf_df['main_kwh'] * blended_rate
                pdf_df['lesco_cost']    = pdf_df['main_kwh'] * lesco_rate
                pdf_df['solar_savings'] = pdf_df['solar_kwh'] * (lesco_rate - blended_rate)
                pdf_buf = generate_pdf(pdf_df, pdf_lts_quick, f"Energy Report — May 2026", blended_rate, lesco_rate)
            st.download_button(
                label="📥 Download PDF",
                data=pdf_buf,
                file_name=f"energy_report_{datetime.now().strftime('%H%M%S')}.pdf",
                mime="application/pdf"
            )

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    '<p style="text-align:center; color:#1e2a42; font-size:0.7rem; letter-spacing:0.1em;">' +
    'SAPPHIRE FIBRES · NEELABLUE DENIM UNIT · ENERGY INTELLIGENCE SYSTEM · v2.0</p>',
    unsafe_allow_html=True
)
