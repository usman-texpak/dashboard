import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Salesperson Margin Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Hide default Streamlit branding */
#MainMenu, footer, header { visibility: hidden; }

/* Main background */
.stApp { background-color: #F8F9FB; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1E2235 0%, #252A40 100%);
    border-right: 1px solid #2E3450;
}
section[data-testid="stSidebar"] * { color: #C8CEDD !important; }
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stMultiSelect label,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #E8EAF2 !important; font-weight: 600; }
section[data-testid="stSidebar"] .stToggle label { color: #C8CEDD !important; }

/* KPI cards */
.kpi-card {
    background: #FFFFFF;
    border-radius: 14px;
    padding: 22px 24px;
    border: 1px solid #E8EBF2;
    box-shadow: 0 2px 12px rgba(0,0,0,0.05);
    height: 100%;
}
.kpi-label {
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #7A8099;
    margin-bottom: 8px;
}
.kpi-value {
    font-size: 28px;
    font-weight: 700;
    color: #1A1F36;
    line-height: 1.1;
    margin-bottom: 4px;
}
.kpi-sub {
    font-size: 12px;
    color: #A0A8C0;
    margin-top: 4px;
}
.kpi-positive { color: #15803D; }
.kpi-negative { color: #DC2626; }
.kpi-neutral  { color: #1A1F36; }

/* Section headers */
.section-header {
    font-size: 17px;
    font-weight: 700;
    color: #1A1F36;
    margin: 8px 0 4px 0;
    padding-bottom: 10px;
    border-bottom: 2px solid #E8EBF2;
}

/* Chart container */
.chart-card {
    background: #FFFFFF;
    border-radius: 14px;
    padding: 20px 24px;
    border: 1px solid #E8EBF2;
    box-shadow: 0 2px 12px rgba(0,0,0,0.05);
    margin-bottom: 18px;
}

/* Tab styling */
.stTabs [data-baseweb="tab-list"] {
    background: #FFFFFF;
    border-radius: 10px;
    padding: 4px;
    border: 1px solid #E8EBF2;
    gap: 2px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px;
    padding: 8px 24px;
    font-weight: 600;
    font-size: 14px;
    color: #7A8099;
}
.stTabs [aria-selected="true"] {
    background: #1E2235 !important;
    color: #FFFFFF !important;
}

/* Status badges */
.badge-above  { background:#DCFCE7; color:#15803D; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:600; }
.badge-within { background:#E0F2FE; color:#0369A1; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:600; }
.badge-below  { background:#FEE2E2; color:#DC2626; padding:3px 10px; border-radius:20px; font-size:12px; font-weight:600; }

/* Summary pill cards */
.pill-row {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin: 12px 0;
}
.pill {
    flex: 1;
    min-width: 140px;
    border-radius: 12px;
    padding: 16px 18px;
    text-align: center;
}
.pill-label { font-size: 11px; font-weight: 700; letter-spacing: 0.07em; text-transform: uppercase; margin-bottom: 6px; }
.pill-val   { font-size: 26px; font-weight: 800; }
.pill-pct   { font-size: 12px; margin-top: 2px; opacity: 0.8; }
.pill-above { background: #DCFCE7; color: #15803D; }
.pill-within{ background: #EFF6FF; color: #1D4ED8; }
.pill-below { background: #FEE2E2; color: #DC2626; }

/* Divider */
.divider { height: 1px; background: #E8EBF2; margin: 18px 0; }

/* Sidebar toggle */
.sidebar-divider {
    height: 1px;
    background: #3A4060;
    margin: 14px 0;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# DATA LOADING & CLEANING
# ─────────────────────────────────────────────
@st.cache_data
def load_data():
    df = pd.read_excel("margin_data.xlsx")

    # Date: primary = Sale Fully Invoiced Date, fallback = Date Validated
    df["_date"] = df["Sale Fully Invoiced Date"].fillna(df["Date Validated"])
    df["Year"] = df["_date"].dt.year

    # Drop rows missing critical fields
    df = df.dropna(subset=["Salesperson"])
    df = df[df["Salesperson"].str.strip() != ""]

    # Numeric coercion
    for col in ["Theoretical Margin", "Real Margin", "Paid SM Amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Paid SM"] = pd.to_numeric(df["Paid SM"], errors="coerce").fillna(0).astype(int)
    df["Paid SM Bool"] = df["Paid SM"] == 1

    # Margin deviation logic – safe division
    theo_abs = df["Theoretical Margin"].abs()
    df["Deviation %"] = np.where(
        theo_abs < 1e-9,
        np.nan,
        (df["Real Margin"] - df["Theoretical Margin"]) / theo_abs,
    )

    # Category
    def categorize(dev):
        if pd.isna(dev):
            return "Within Range (±5%)"
        if dev > 0.05:
            return "Above 5%"
        if dev < -0.05:
            return "Less than 5%"
        return "Within Range (±5%)"

    df["Margin Category"] = df["Deviation %"].apply(categorize)
    df["Year"] = df["Year"].astype("Int64")

    return df


df_all = load_data()

CAT_ORDER  = ["Above 5%", "Within Range (±5%)", "Less than 5%"]
CAT_COLORS = {"Above 5%": "#16A34A", "Within Range (±5%)": "#3B82F6", "Less than 5%": "#DC2626"}
CAT_LIGHT  = {"Above 5%": "#DCFCE7", "Within Range (±5%)": "#EFF6FF", "Less than 5%": "#FEE2E2"}
CAT_TEXT   = {"Above 5%": "#15803D", "Within Range (±5%)": "#1D4ED8", "Less than 5%": "#DC2626"}

CHART_LAYOUT = dict(
    font_family="Inter",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=10, r=10, t=36, b=10),
)


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Dashboard Controls")
    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

    available_years = sorted(df_all["Year"].dropna().unique().tolist(), reverse=True)
    selected_years = st.multiselect(
        "📅 Select Year(s)",
        options=available_years,
        default=available_years[:2] if len(available_years) >= 2 else available_years,
        help="Filter all charts and tables by invoice year",
    )

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
    only_paid_sm = st.toggle(
        "💬 Show Only SMS Paid Projects",
        value=False,
        help="Filter to projects where Paid SM = 1",
    )

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
    st.markdown("### 🎯 Category Filter")
    cat_filter = st.multiselect(
        "Margin Categories",
        options=CAT_ORDER,
        default=CAT_ORDER,
    )

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:11px;color:#5A6080;line-height:1.6'>"
        "Margin deviation = (Real − Theoretical) / |Theoretical|<br>"
        "5% threshold applied strictly per project."
        "</div>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# FILTERED DATA
# ─────────────────────────────────────────────
df = df_all.copy()
if selected_years:
    df = df[df["Year"].isin(selected_years)]
if only_paid_sm:
    df = df[df["Paid SM Bool"]]
if cat_filter:
    df = df[df["Margin Category"].isin(cat_filter)]


# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────
def fmt_euro(v):
    if abs(v) >= 1_000_000:
        return f"€{v/1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"€{v/1_000:.1f}K"
    return f"€{v:,.0f}"

def kpi_card(label, value, sub="", color_class="kpi-neutral"):
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {color_class}">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""

def pill_html(cat, count, total):
    pct = f"{count/total*100:.1f}%" if total > 0 else "—"
    key = "above" if "Above" in cat else ("within" if "Within" in cat else "below")
    short = "Above 5%" if key == "above" else ("Within ±5%" if key == "within" else "Below −5%")
    return f"""
    <div class="pill pill-{key}">
        <div class="pill-label">{short}</div>
        <div class="pill-val">{count}</div>
        <div class="pill-pct">{pct} of total</div>
    </div>"""


# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div style="padding: 8px 0 22px 0;">
    <div style="font-size:26px;font-weight:800;color:#1A1F36;letter-spacing:-0.3px;">
        Salesperson Performance Dashboard
    </div>
    <div style="font-size:14px;color:#7A8099;margin-top:4px;">
        Margin accuracy analysis · 5% deviation threshold · SMS commission tracking
    </div>
</div>
""", unsafe_allow_html=True)

if not selected_years:
    st.warning("⚠️ Please select at least one year from the sidebar.")
    st.stop()

year_label = ", ".join(str(y) for y in sorted(selected_years))
st.markdown(
    f"<div style='display:inline-block;background:#1E2235;color:#C8CEDD;padding:5px 14px;"
    f"border-radius:20px;font-size:12px;font-weight:600;margin-bottom:16px;'>"
    f"📅 Viewing: {year_label}{'  ·  💬 SMS Paid Only' if only_paid_sm else ''}</div>",
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2 = st.tabs(["🌐  Global Team Statistics", "🔍  Salesperson Deep-Dive"])


# ══════════════════════════════════════════════
# TAB 1: GLOBAL TEAM STATISTICS
# ══════════════════════════════════════════════
with tab1:

    # ── KPI ROW ──
    total_theo  = df["Theoretical Margin"].sum()
    total_real  = df["Real Margin"].sum()
    variance    = ((total_real - total_theo) / abs(total_theo) * 100) if abs(total_theo) > 1e-9 else 0
    total_proj  = len(df)
    paid_sm_cnt = df["Paid SM Bool"].sum()
    paid_sm_amt = df.loc[df["Paid SM Bool"], "Paid SM Amount"].sum()

    var_class = "kpi-positive" if variance >= 0 else "kpi-negative"
    var_sign  = "+" if variance >= 0 else ""

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(kpi_card("Total Theoretical Margin", fmt_euro(total_theo), f"Across {total_proj:,} projects"), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card("Total Real Margin", fmt_euro(total_real), "Realised margin"), unsafe_allow_html=True)
    with c3:
        st.markdown(kpi_card("Overall Variance", f"{var_sign}{variance:.1f}%", "vs. theoretical", var_class), unsafe_allow_html=True)
    with c4:
        st.markdown(kpi_card("Total Projects", f"{total_proj:,}", f"{len(df['Salesperson'].unique())} salespersons"), unsafe_allow_html=True)
    with c5:
        st.markdown(kpi_card("SMS Paid Projects", f"{paid_sm_cnt:,}", f"Total: {fmt_euro(paid_sm_amt)}"), unsafe_allow_html=True)

    st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

    # ── CATEGORY SUMMARY PILLS ──
    above_n  = (df["Margin Category"] == "Above 5%").sum()
    within_n = (df["Margin Category"] == "Within Range (±5%)").sum()
    below_n  = (df["Margin Category"] == "Less than 5%").sum()

    st.markdown(
        f"""<div class="pill-row">
            {pill_html("Above 5%", above_n, total_proj)}
            {pill_html("Within Range (±5%)", within_n, total_proj)}
            {pill_html("Less than 5%", below_n, total_proj)}
        </div>""",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

    # ── STACKED BAR: Team standings ──
    col_left, col_right = st.columns([3, 2], gap="large")

    with col_left:
        st.markdown("<div class='section-header'>Team Standings · Project Distribution by Category</div>", unsafe_allow_html=True)

        sp_cat = (
            df.groupby(["Salesperson", "Margin Category"])
            .size()
            .reset_index(name="Count")
        )
        sp_total = df.groupby("Salesperson").size().reset_index(name="Total")
        sp_cat = sp_cat.merge(sp_total, on="Salesperson")
        sp_cat["Pct"] = (sp_cat["Count"] / sp_cat["Total"] * 100).round(1)

        sp_order = (
            df[df["Margin Category"] == "Above 5%"]
            .groupby("Salesperson")
            .size()
            .sort_values(ascending=True)
            .index.tolist()
        )
        all_sp = df["Salesperson"].unique().tolist()
        sp_order = [s for s in sp_order if s in all_sp] + [s for s in all_sp if s not in sp_order]

        fig_bar = go.Figure()
        for cat in CAT_ORDER:
            sub = sp_cat[sp_cat["Margin Category"] == cat]
            sub = sub.set_index("Salesperson").reindex(sp_order).reset_index()
            fig_bar.add_trace(go.Bar(
                name=cat,
                y=sub["Salesperson"],
                x=sub["Count"].fillna(0),
                orientation="h",
                marker_color=CAT_COLORS[cat],
                text=sub["Count"].fillna(0).astype(int),
                textposition="inside",
                insidetextanchor="middle",
                customdata=sub["Pct"].fillna(0),
                hovertemplate="<b>%{y}</b><br>" + cat + ": %{x} projects (%{customdata:.1f}%)<extra></extra>",
            ))

        fig_bar.update_layout(
            **CHART_LAYOUT,
            barmode="stack",
            height=max(380, len(sp_order) * 34 + 60),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                        font_size=12, itemsizing="constant"),
            xaxis=dict(title="Number of Projects", gridcolor="#F0F1F5", tickfont_size=11),
            yaxis=dict(tickfont_size=11, automargin=True),
        )
        st.plotly_chart(fig_bar, use_container_width=True, config={"displayModeBar": False})

    with col_right:
        st.markdown("<div class='section-header'>Margin Accuracy · Year-over-Year Trend</div>", unsafe_allow_html=True)

        yoy = (
            df_all[df_all["Year"].notna()]
            .groupby(["Year", "Margin Category"])
            .size()
            .reset_index(name="Count")
        )
        yoy_total = df_all[df_all["Year"].notna()].groupby("Year").size().reset_index(name="Total")
        yoy = yoy.merge(yoy_total, on="Year")
        yoy["Pct"] = (yoy["Count"] / yoy["Total"] * 100).round(1)

        fig_yoy = px.bar(
            yoy,
            x="Year",
            y="Pct",
            color="Margin Category",
            color_discrete_map=CAT_COLORS,
            barmode="stack",
            text="Pct",
            category_orders={"Margin Category": CAT_ORDER},
        )
        fig_yoy.update_traces(texttemplate="%{text:.0f}%", textposition="inside")
        fig_yoy.update_layout(
            **CHART_LAYOUT,
            height=280,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, font_size=11),
            xaxis=dict(title="Year", type="category", tickfont_size=11),
            yaxis=dict(title="% of Projects", range=[0, 102], gridcolor="#F0F1F5", tickfont_size=11),
            showlegend=True,
        )
        st.plotly_chart(fig_yoy, use_container_width=True, config={"displayModeBar": False})

        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
        st.markdown("<div class='section-header'>SMS Paid vs Unpaid by Margin Category</div>", unsafe_allow_html=True)

        sms_cross = (
            df.groupby(["Margin Category", "Paid SM Bool"])
            .size()
            .reset_index(name="Count")
        )
        sms_cross["SMS Status"] = sms_cross["Paid SM Bool"].map({True: "SMS Paid", False: "SMS Unpaid"})

        fig_sms = px.bar(
            sms_cross,
            x="Margin Category",
            y="Count",
            color="SMS Status",
            color_discrete_map={"SMS Paid": "#7C3AED", "SMS Unpaid": "#C4B5FD"},
            barmode="group",
            category_orders={"Margin Category": CAT_ORDER},
        )
        fig_sms.update_layout(
            **CHART_LAYOUT,
            height=240,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, font_size=11),
            xaxis=dict(tickfont_size=11, title=""),
            yaxis=dict(title="Projects", gridcolor="#F0F1F5", tickfont_size=11),
        )
        st.plotly_chart(fig_sms, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

    # ── SCATTER: Theoretical vs Real Margin ──
    st.markdown("<div class='section-header'>Theoretical vs. Real Margin · Coloured by SMS Paid Status</div>", unsafe_allow_html=True)

    df_scatter = df[df["Theoretical Margin"].notna() & df["Real Margin"].notna()].copy()
    df_scatter["SMS Status"] = df_scatter["Paid SM Bool"].map({True: "SMS Paid ✓", False: "SMS Unpaid"})
    df_scatter["Deviation Label"] = df_scatter["Deviation %"].apply(
        lambda x: f"{x*100:+.1f}%" if pd.notna(x) else "N/A"
    )

    clip_val = df_scatter[["Theoretical Margin", "Real Margin"]].abs().quantile(0.98).max()
    df_scatter_vis = df_scatter[
        (df_scatter["Theoretical Margin"].abs() <= clip_val) &
        (df_scatter["Real Margin"].abs() <= clip_val)
    ]

    fig_scatter = px.scatter(
        df_scatter_vis,
        x="Theoretical Margin",
        y="Real Margin",
        color="SMS Status",
        color_discrete_map={"SMS Paid ✓": "#7C3AED", "SMS Unpaid": "#94A3B8"},
        symbol="Margin Category",
        symbol_map={"Above 5%": "circle", "Within Range (±5%)": "square", "Less than 5%": "x"},
        opacity=0.72,
        hover_data={
            "Analytic Account": True,
            "Customer": True,
            "Salesperson": True,
            "Deviation Label": True,
            "Theoretical Margin": ":.1f",
            "Real Margin": ":.1f",
        },
        labels={"Theoretical Margin": "Theoretical Margin (€)", "Real Margin": "Real Margin (€)"},
    )

    # Perfect accuracy line
    mn = min(df_scatter_vis["Theoretical Margin"].min(), df_scatter_vis["Real Margin"].min())
    mx = max(df_scatter_vis["Theoretical Margin"].max(), df_scatter_vis["Real Margin"].max())
    fig_scatter.add_trace(go.Scatter(
        x=[mn, mx], y=[mn, mx],
        mode="lines",
        name="Perfect Accuracy",
        line=dict(color="#9CA3AF", dash="dash", width=1.5),
        hoverinfo="skip",
    ))
    fig_scatter.add_trace(go.Scatter(
        x=[mn, mx], y=[mn * 1.05, mx * 1.05],
        mode="lines",
        name="+5% Band",
        line=dict(color="#86EFAC", dash="dot", width=1),
        hoverinfo="skip",
    ))
    fig_scatter.add_trace(go.Scatter(
        x=[mn, mx], y=[mn * 0.95, mx * 0.95],
        mode="lines",
        name="−5% Band",
        line=dict(color="#FCA5A5", dash="dot", width=1),
        hoverinfo="skip",
    ))

    fig_scatter.update_layout(
        **CHART_LAYOUT,
        height=460,
        legend=dict(orientation="h", yanchor="bottom", y=1.01, font_size=12),
        xaxis=dict(gridcolor="#F0F1F5", zerolinecolor="#E0E3EC", tickfont_size=11),
        yaxis=dict(gridcolor="#F0F1F5", zerolinecolor="#E0E3EC", tickfont_size=11),
    )
    st.plotly_chart(fig_scatter, use_container_width=True, config={"displayModeBar": True})


# ══════════════════════════════════════════════
# TAB 2: SALESPERSON DEEP-DIVE
# ══════════════════════════════════════════════
with tab2:

    all_salespersons = sorted(df["Salesperson"].unique())

    col_pick, col_year = st.columns([2, 1])
    with col_pick:
        selected_sp = st.selectbox(
            "👤 Select Salesperson",
            options=all_salespersons,
            index=0,
        )
    with col_year:
        sp_year_options = ["All Years"] + [str(y) for y in sorted(selected_years, reverse=True)]
        sp_year = st.selectbox("🗓 Focus Year", options=sp_year_options, index=0)

    df_sp = df[df["Salesperson"] == selected_sp].copy()
    if sp_year != "All Years":
        df_sp = df_sp[df_sp["Year"] == int(sp_year)]

    if df_sp.empty:
        st.info("No projects found for this salesperson with the current filters.")
        st.stop()

    # ── PERFORMANCE SUMMARY PILLS ──
    sp_above  = (df_sp["Margin Category"] == "Above 5%").sum()
    sp_within = (df_sp["Margin Category"] == "Within Range (±5%)").sum()
    sp_below  = (df_sp["Margin Category"] == "Less than 5%").sum()
    sp_total  = len(df_sp)
    sp_theo   = df_sp["Theoretical Margin"].sum()
    sp_real   = df_sp["Real Margin"].sum()
    sp_var    = ((sp_real - sp_theo) / abs(sp_theo) * 100) if abs(sp_theo) > 1e-9 else 0
    sp_paid_sm = df_sp["Paid SM Bool"].sum()

    st.markdown(f"""
    <div style="background:#FFFFFF;border-radius:14px;padding:20px 24px;
                border:1px solid #E8EBF2;box-shadow:0 2px 12px rgba(0,0,0,0.05);margin-bottom:18px;">
        <div style="font-size:18px;font-weight:700;color:#1A1F36;margin-bottom:14px;">
            {selected_sp}
            <span style="font-size:13px;font-weight:500;color:#7A8099;margin-left:8px;">
                — {sp_total} projects · {sp_year}
            </span>
        </div>
        <div class="pill-row">
            {pill_html("Above 5%", sp_above, sp_total)}
            {pill_html("Within Range (±5%)", sp_within, sp_total)}
            {pill_html("Less than 5%", sp_below, sp_total)}
        </div>
        <div style="display:flex;gap:32px;margin-top:14px;flex-wrap:wrap;">
            <div><span style="font-size:11px;color:#7A8099;font-weight:600;text-transform:uppercase;letter-spacing:0.06em">Theoretical</span>
                 <div style="font-size:18px;font-weight:700;color:#1A1F36">{fmt_euro(sp_theo)}</div></div>
            <div><span style="font-size:11px;color:#7A8099;font-weight:600;text-transform:uppercase;letter-spacing:0.06em">Real Margin</span>
                 <div style="font-size:18px;font-weight:700;color:#1A1F36">{fmt_euro(sp_real)}</div></div>
            <div><span style="font-size:11px;color:#7A8099;font-weight:600;text-transform:uppercase;letter-spacing:0.06em">Overall Variance</span>
                 <div style="font-size:18px;font-weight:700;color:{'#15803D' if sp_var >= 0 else '#DC2626'}">{'+' if sp_var >= 0 else ''}{sp_var:.1f}%</div></div>
            <div><span style="font-size:11px;color:#7A8099;font-weight:600;text-transform:uppercase;letter-spacing:0.06em">SMS Paid</span>
                 <div style="font-size:18px;font-weight:700;color:#7C3AED">{sp_paid_sm} projects</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── MINI CHARTS ROW ──
    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown("<div class='section-header'>Project Category Breakdown</div>", unsafe_allow_html=True)
        fig_pie = go.Figure(go.Pie(
            labels=CAT_ORDER,
            values=[sp_above, sp_within, sp_below],
            hole=0.55,
            marker_colors=[CAT_COLORS[c] for c in CAT_ORDER],
            textinfo="label+percent",
            textfont_size=12,
            hovertemplate="%{label}: %{value} projects (%{percent})<extra></extra>",
        ))
        fig_pie.add_annotation(
            text=f"<b>{sp_total}</b><br>projects",
            x=0.5, y=0.5, font_size=16, showarrow=False,
            font_color="#1A1F36",
        )
        fig_pie.update_layout(**CHART_LAYOUT, height=280, showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True, config={"displayModeBar": False})

    with c2:
        st.markdown("<div class='section-header'>Year-by-Year Category Trend</div>", unsafe_allow_html=True)
        sp_yoy = (
            df_all[df_all["Salesperson"] == selected_sp]
            .groupby(["Year", "Margin Category"])
            .size()
            .reset_index(name="Count")
        )
        fig_sp_yoy = px.bar(
            sp_yoy,
            x="Year",
            y="Count",
            color="Margin Category",
            color_discrete_map=CAT_COLORS,
            barmode="stack",
            category_orders={"Margin Category": CAT_ORDER},
        )
        fig_sp_yoy.update_layout(
            **CHART_LAYOUT,
            height=280,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, font_size=11),
            xaxis=dict(type="category", tickfont_size=11, title=""),
            yaxis=dict(title="Projects", gridcolor="#F0F1F5", tickfont_size=11),
        )
        st.plotly_chart(fig_sp_yoy, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

    # ── PROJECT DETAIL TABLE ──
    st.markdown("<div class='section-header'>Project Detail Table</div>", unsafe_allow_html=True)

    search_q = st.text_input("🔍 Search projects", placeholder="Filter by account, customer…", label_visibility="collapsed")

    display_cols = [
        "Analytic Account", "Customer", "State",
        "Theoretical Margin", "Real Margin",
        "Deviation %", "Margin Category",
        "Paid SM Bool", "Paid SM Amount",
        "Year",
    ]
    df_table = df_sp[display_cols].copy()
    df_table["Deviation %"] = (df_table["Deviation %"] * 100).round(2)
    df_table["Paid SM"] = df_table["Paid SM Bool"].map({True: "✅ Paid", False: "—"})
    df_table.drop(columns=["Paid SM Bool"], inplace=True)
    df_table.rename(columns={
        "Theoretical Margin": "Theoretical (€)",
        "Real Margin": "Real (€)",
        "Deviation %": "Deviation (%)",
        "Paid SM Amount": "SMS Amount (€)",
    }, inplace=True)

    if search_q:
        mask = df_table["Analytic Account"].str.contains(search_q, case=False, na=False) | \
               df_table["Customer"].str.contains(search_q, case=False, na=False)
        df_table = df_table[mask]

    def highlight_row(row):
        cat = row.get("Margin Category", "")
        if cat == "Above 5%":
            bg = "#F0FDF4"
            color = "#15803D"
        elif cat == "Less than 5%":
            bg = "#FFF1F2"
            color = "#DC2626"
        else:
            bg = "#EFF6FF"
            color = "#1D4ED8"
        return [f"background-color:{bg};color:{color};font-weight:500" if c == "Margin Category"
                else f"background-color:{bg}" for c in row.index]

    def style_deviation(val):
        try:
            v = float(val)
            if v > 5:   return "color:#15803D;font-weight:600"
            if v < -5:  return "color:#DC2626;font-weight:600"
            return "color:#1D4ED8;font-weight:600"
        except Exception:
            return ""

    styled = (
        df_table.style
        .apply(highlight_row, axis=1)
        .applymap(style_deviation, subset=["Deviation (%)"])
        .format({
            "Theoretical (€)": "€{:,.2f}",
            "Real (€)": "€{:,.2f}",
            "SMS Amount (€)": "€{:,.2f}",
            "Deviation (%)": "{:+.2f}%",
        }, na_rep="—")
    )

    st.dataframe(styled, use_container_width=True, height=min(600, max(300, len(df_table) * 36 + 50)))
    st.caption(f"Showing {len(df_table):,} projects · Green = Above 5% · Blue = Within ±5% · Red = Less than −5%")
