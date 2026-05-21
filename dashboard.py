import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Salesperson Performance", page_icon="📊",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
#MainMenu,footer,header{visibility:hidden;}
.stApp{background:#F0F2F8;}
section[data-testid="stSidebar"]{background:#1A1F36;}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] span{color:#A0AABF !important;}
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3{color:#fff !important;}
.block-container{padding:2rem 2.5rem;}
.stTabs [data-baseweb="tab-list"]{background:#fff;border-radius:12px;padding:5px;border:1px solid #E2E8F0;gap:4px;}
.stTabs [data-baseweb="tab"]{border-radius:9px;padding:9px 28px;font-weight:600;font-size:14px;color:#64748B;}
.stTabs [aria-selected="true"]{background:#1A1F36 !important;color:#fff !important;}
div[data-testid="stMetric"]{background:white;border-radius:14px;padding:20px;border:1px solid #E8EDF5;box-shadow:0 2px 8px rgba(0,0,0,0.06);}
div[data-testid="stMetricValue"]{font-size:28px !important;font-weight:800 !important;color:#1A1F36 !important;}
div[data-testid="stMetricLabel"]{font-size:13px !important;font-weight:600 !important;color:#64748B !important;}
div[data-testid="stMetricDelta"]{font-size:12px !important;}
div[data-testid="stAlert"]{border-radius:14px !important;padding:20px 24px !important;border:none !important;box-shadow:0 2px 8px rgba(0,0,0,0.06);}
div[data-testid="stAlert"] p{font-size:15px !important;}
</style>""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    df = pd.read_excel("margin_data.xlsx")
    df["_date"] = df["Sale Fully Invoiced Date"].fillna(df["Date Validated"])
    df["Year"] = df["_date"].dt.year
    df = df.dropna(subset=["Salesperson"])
    df = df[df["Salesperson"].str.strip() != ""]
    for col in ["Theoretical Margin","Real Margin","Paid SM Amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Paid SM"] = pd.to_numeric(df["Paid SM"], errors="coerce").fillna(0).astype(int)
    df["Paid SM Bool"] = df["Paid SM"] == 1
    theo_abs = df["Theoretical Margin"].abs()
    df["Deviation"] = np.where(theo_abs < 1e-9, np.nan,
        (df["Real Margin"] - df["Theoretical Margin"]) / theo_abs)
    def cat(d):
        if pd.isna(d): return "Within Range (±5%)"
        if d > 0.05:   return "Above 5%"
        if d < -0.05:  return "Less than 5%"
        return "Within Range (±5%)"
    df["Category"] = df["Deviation"].apply(cat)
    df["Year"] = df["Year"].astype("Int64")
    return df

# ── LOGIN ──────────────────────────────────────
USERS = {
    "admin":   "texpak2024",
    "usman":   "dashboard123",
    "manager": "margin2024",
}

def check_login():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if "username" not in st.session_state:
        st.session_state.username = ""

check_login()

if not st.session_state.logged_in:
    st.markdown("<div style='height:60px'></div>", unsafe_allow_html=True)
    col_l, col_m, col_r = st.columns([1, 1.2, 1])
    with col_m:
        st.markdown("""
        <div style='background:white;border-radius:20px;padding:40px 36px;
                    box-shadow:0 8px 32px rgba(0,0,0,0.12);text-align:center;
                    border:1px solid #E2E8F0'>
            <div style='font-size:48px;margin-bottom:8px'>📊</div>
            <h2 style='font-size:22px;font-weight:800;color:#1A1F36;margin:0 0 6px 0'>
                Salesperson Dashboard</h2>
            <p style='color:#64748B;font-size:14px;margin:0 0 28px 0'>
                Please sign in to continue</p>
        </div>""", unsafe_allow_html=True)

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        username = st.text_input("👤 Username", placeholder="Enter username")
        password = st.text_input("🔑 Password", type="password", placeholder="Enter password")
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

        if st.button("🔓 Sign In", use_container_width=True, type="primary"):
            if username in USERS and USERS[username] == password:
                st.session_state.logged_in = True
                st.session_state.username  = username
                st.rerun()
            else:
                st.error("❌ Incorrect username or password. Please try again.")

        st.markdown("""
        <div style='margin-top:16px;padding:12px;background:#F8FAFC;border-radius:10px;
                    border:1px solid #E2E8F0;font-size:12px;color:#64748B;text-align:center'>
            🔒 Secure access · Contact admin for credentials
        </div>""", unsafe_allow_html=True)
    st.stop()

# ── FILE UPLOAD & DATA LOADING ───────────────
@st.cache_data
def process_data(file_bytes):
    import io
    df = pd.read_excel(io.BytesIO(file_bytes))
    df["_date"] = df["Sale Fully Invoiced Date"].fillna(df["Date Validated"])
    df["Year"] = df["_date"].dt.year
    df = df.dropna(subset=["Salesperson"])
    df = df[df["Salesperson"].str.strip() != ""]
    for col in ["Theoretical Margin","Real Margin","Paid SM Amount"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Paid SM"] = pd.to_numeric(df["Paid SM"], errors="coerce").fillna(0).astype(int)
    df["Paid SM Bool"] = df["Paid SM"] == 1
    theo_abs = df["Theoretical Margin"].abs()
    df["Deviation"] = np.where(theo_abs < 1e-9, np.nan,
        (df["Real Margin"] - df["Theoretical Margin"]) / theo_abs)
    def cat(d):
        if pd.isna(d): return "Within Range (±5%)"
        if d > 0.05:   return "Above 5%"
        if d < -0.05:  return "Less than 5%"
        return "Within Range (±5%)"
    df["Category"] = df["Deviation"].apply(cat)
    df["Year"] = df["Year"].astype("Int64")
    return df

# Use uploaded file if available, else use default
if "uploaded_bytes" in st.session_state and st.session_state.uploaded_bytes is not None:
    df_all = process_data(st.session_state.uploaded_bytes)
else:
    df_all = load_data()
COLORS    = {"Above 5%":"#16A34A","Within Range (±5%)":"#2563EB","Less than 5%":"#DC2626"}
CAT_ORDER = ["Above 5%","Within Range (±5%)","Less than 5%"]
CHART_CFG = dict(font_family="Inter",paper_bgcolor="rgba(0,0,0,0)",
                 plot_bgcolor="rgba(0,0,0,0)",margin=dict(l=8,r=8,t=32,b=8))

def fmt(v):
    if abs(v)>=1_000_000: return f"€{v/1_000_000:.2f}M"
    if abs(v)>=1_000:     return f"€{v/1_000:.1f}K"
    return f"€{v:,.0f}"

# ── SIDEBAR ──
with st.sidebar:
    # User info + logout
    st.markdown(f"""
    <div style='background:#252D47;border-radius:12px;padding:14px 16px;margin-bottom:4px'>
        <div style='font-size:11px;color:#6B7FA3;font-weight:600;text-transform:uppercase;
                    letter-spacing:.06em;margin-bottom:4px'>Logged in as</div>
        <div style='font-size:15px;font-weight:700;color:#FFFFFF'>
            👤 {st.session_state.username}</div>
    </div>""", unsafe_allow_html=True)
    if st.button("🚪 Sign Out", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username  = ""
        st.rerun()

    st.markdown("---")
    st.markdown("## 📊 Dashboard Controls")
    st.markdown("---")

    # File upload
    st.markdown("### 📤 Update Data")
    uploaded_file = st.file_uploader(
        "Upload new Excel file",
        type=["xlsx","xls"],
        help="Upload a new version of the data file to refresh the dashboard",
        label_visibility="collapsed"
    )
    if uploaded_file is not None:
        file_bytes = uploaded_file.read()
        if ("uploaded_bytes" not in st.session_state or
                st.session_state.uploaded_bytes != file_bytes):
            st.session_state.uploaded_bytes = file_bytes
            process_data.clear()
            st.success(f"✅ {uploaded_file.name} loaded!")
            st.rerun()
    else:
        if "uploaded_bytes" not in st.session_state:
            st.session_state.uploaded_bytes = None
        st.caption("📁 Using default data file")

    if st.session_state.get("uploaded_bytes"):
        if st.button("🔄 Reset to Default Data", use_container_width=True):
            st.session_state.uploaded_bytes = None
            process_data.clear()
            st.rerun()

    st.markdown("---")
    years_avail = sorted(df_all["Year"].dropna().unique().tolist(), reverse=True)
    sel_years = st.multiselect("📅 Year(s)", years_avail,
                               default=years_avail[:2] if len(years_avail)>=2 else years_avail)
    st.markdown("---")
    only_sms = st.toggle("💬 SMS Paid Only", value=False)
    st.markdown("---")
    sel_cats = st.multiselect("🎯 Categories", CAT_ORDER, default=CAT_ORDER)
    st.markdown("---")

    # Account State filter
    st.markdown("### 📂 Account Status")
    state_options = ["Open", "Validated", "Submitted", "Cancelled"]
    sel_states = st.multiselect(
        "Account State",
        options=state_options,
        default=state_options,
        help="Open = active · Validated = closed & approved · Submitted = pending · Cancelled",
        label_visibility="collapsed"
    )

    st.markdown("---")

    # Payment status filter
    st.markdown("### 💳 Payment Status")
    payment_options = ["✅ Fully Paid", "📄 Invoiced (Not Paid)", "⏳ Not Yet Invoiced"]
    sel_payment = st.multiselect(
        "Payment Status",
        options=payment_options,
        default=payment_options,
        label_visibility="collapsed",
        help="Paid = Sale Fully Paid Date exists · Invoiced = Sale Fully Invoiced Date exists"
    )

    st.markdown("---")
    st.caption("Deviation = (Real − Theo) / |Theo| · 5% threshold per project")

df = df_all.copy()
if sel_years:   df = df[df["Year"].isin(sel_years)]
if only_sms:    df = df[df["Paid SM Bool"]]
if sel_cats:    df = df[df["Category"].isin(sel_cats)]

# State filter
if sel_states:  df = df[df["State"].isin(sel_states)]

# Payment status filter
def get_payment_status(row):
    if pd.notna(row["Sale Fully Paid Date"]):
        return "✅ Fully Paid"
    elif pd.notna(row["Sale Fully Invoiced Date"]):
        return "📄 Invoiced (Not Paid)"
    else:
        return "⏳ Not Yet Invoiced"

df["Payment Status"] = df.apply(get_payment_status, axis=1)
df_all["Payment Status"] = df_all.apply(get_payment_status, axis=1)

if sel_payment: df = df[df["Payment Status"].isin(sel_payment)]

if not sel_years:
    st.warning("⚠️ Select at least one year from the sidebar.")
    st.stop()

# ── HEADER ──
st.markdown(f"# 📊 Salesperson Performance Dashboard")
st.markdown(f"Margin accuracy · 5% deviation threshold · SMS commission tracking")
yr_label = ", ".join(str(y) for y in sorted(sel_years))
st.info(f"📅 Viewing: {yr_label}{'  ·  💬 SMS Paid Only' if only_sms else ''}")
st.markdown("---")

tab1, tab2 = st.tabs(["🌐  Global Team Statistics", "🔍  Salesperson Deep-Dive"])

# ══ TAB 1 ══
with tab1:
    total_theo = df["Theoretical Margin"].sum()
    total_real = df["Real Margin"].sum()
    variance   = ((total_real-total_theo)/abs(total_theo)*100) if abs(total_theo)>1e-9 else 0
    total_proj = len(df)
    sms_count  = df["Paid SM Bool"].sum()
    sms_amt    = df.loc[df["Paid SM Bool"],"Paid SM Amount"].sum()
    above_n    = (df["Category"]=="Above 5%").sum()
    within_n   = (df["Category"]=="Within Range (±5%)").sum()
    below_n    = (df["Category"]=="Less than 5%").sum()

    # KPI cards
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("💰 Theoretical", fmt(total_theo), f"{total_proj:,} projects")
    k2.metric("📈 Real Margin", fmt(total_real), "Realised")
    k3.metric("📊 Variance",    f"{'+' if variance>=0 else ''}{variance:.1f}%", "vs theoretical")
    k4.metric("📁 Projects",    f"{total_proj:,}", f"{df['Salesperson'].nunique()} people")
    k5.metric("💬 SMS Paid",    f"{sms_count:,}", fmt(sms_amt))

    st.markdown("<br>", unsafe_allow_html=True)

    # Category cards using columns + native st components
    pct_a = above_n/total_proj*100 if total_proj else 0
    pct_b = within_n/total_proj*100 if total_proj else 0
    pct_c = below_n/total_proj*100 if total_proj else 0

    ca, cb, cc = st.columns(3)
    with ca:
        st.markdown(f'''<div style="background:#D1FAE5;border-radius:16px;padding:24px 28px;border-left:6px solid #16A34A;box-shadow:0 2px 8px rgba(0,0,0,0.07)">
        <p style="font-size:12px;font-weight:700;color:#15803D;text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px 0">🟢 ABOVE 5%</p>
        <p style="font-size:44px;font-weight:800;color:#14532D;margin:0;line-height:1">{above_n:,}</p>
        <p style="font-size:14px;color:#166534;margin:8px 0 0 0;font-weight:500">{pct_a:.1f}% of total projects</p></div>''', unsafe_allow_html=True)
    with cb:
        st.markdown(f'''<div style="background:#DBEAFE;border-radius:16px;padding:24px 28px;border-left:6px solid #2563EB;box-shadow:0 2px 8px rgba(0,0,0,0.07)">
        <p style="font-size:12px;font-weight:700;color:#1D4ED8;text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px 0">🔵 WITHIN ±5%</p>
        <p style="font-size:44px;font-weight:800;color:#1E3A8A;margin:0;line-height:1">{within_n:,}</p>
        <p style="font-size:14px;color:#1E40AF;margin:8px 0 0 0;font-weight:500">{pct_b:.1f}% of total projects</p></div>''', unsafe_allow_html=True)
    with cc:
        st.markdown(f'''<div style="background:#FEE2E2;border-radius:16px;padding:24px 28px;border-left:6px solid #DC2626;box-shadow:0 2px 8px rgba(0,0,0,0.07)">
        <p style="font-size:12px;font-weight:700;color:#DC2626;text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px 0">🔴 LESS THAN 5%</p>
        <p style="font-size:44px;font-weight:800;color:#7F1D1D;margin:0;line-height:1">{below_n:,}</p>
        <p style="font-size:14px;color:#991B1B;margin:8px 0 0 0;font-weight:500">{pct_c:.1f}% of total projects</p></div>''', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Payment status summary
    paid_n     = (df["Payment Status"] == "✅ Fully Paid").sum()
    invoiced_n = (df["Payment Status"] == "📄 Invoiced (Not Paid)").sum()
    pending_n  = (df["Payment Status"] == "⏳ Not Yet Invoiced").sum()

    st.markdown("#### 💳 Payment Status Overview")
    p1, p2, p3, p4, p5 = st.columns(5)
    p1.metric("✅ Fully Paid",        f"{paid_n:,}",     f"{paid_n/total_proj*100:.1f}%")
    p2.metric("📄 Invoiced (Not Paid)", f"{invoiced_n:,}", f"{invoiced_n/total_proj*100:.1f}%")
    p3.metric("⏳ Not Yet Invoiced",   f"{pending_n:,}",  f"{pending_n/total_proj*100:.1f}%")

    # Account state summary
    st.markdown("#### 📂 Account State Overview")
    s1, s2, s3, s4 = st.columns(4)
    open_n      = (df["State"] == "Open").sum()
    validated_n = (df["State"] == "Validated").sum()
    submitted_n = (df["State"] == "Submitted").sum()
    cancelled_n = (df["State"] == "Cancelled").sum()
    s1.metric("🟢 Open",       f"{open_n:,}",      f"{open_n/total_proj*100:.1f}% of total")
    s2.metric("✅ Validated",  f"{validated_n:,}",  f"{validated_n/total_proj*100:.1f}% of total")
    s3.metric("📤 Submitted",  f"{submitted_n:,}",  f"{submitted_n/total_proj*100:.1f}% of total")
    s4.metric("❌ Cancelled",  f"{cancelled_n:,}",  f"{cancelled_n/total_proj*100:.1f}% of total")

    st.markdown("---")

    # Charts row
    left, right = st.columns([3,2], gap="large")

    with left:
        st.markdown("#### Team standings · Projects by margin category")
        sp_cat = df.groupby(["Salesperson","Category"]).size().reset_index(name="Count")
        sp_tot = df.groupby("Salesperson").size().reset_index(name="Total")
        sp_cat = sp_cat.merge(sp_tot, on="Salesperson")
        sp_order = (df[df["Category"]=="Above 5%"].groupby("Salesperson").size()
                    .sort_values(ascending=True).index.tolist())
        all_sp = df["Salesperson"].unique().tolist()
        sp_order = [s for s in sp_order if s in all_sp]+[s for s in all_sp if s not in sp_order]
        fig = go.Figure()
        for c in CAT_ORDER:
            sub = sp_cat[sp_cat["Category"]==c].set_index("Salesperson").reindex(sp_order).reset_index()
            fig.add_trace(go.Bar(name=c, y=sub["Salesperson"], x=sub["Count"].fillna(0),
                orientation="h", marker_color=COLORS[c],
                text=sub["Count"].fillna(0).astype(int),
                textposition="inside", insidetextanchor="middle",
                hovertemplate="<b>%{y}</b><br>"+c+": %{x}<extra></extra>"))
        fig.update_layout(**CHART_CFG, barmode="stack",
                          height=max(360,len(sp_order)*34+60),
                          legend=dict(orientation="h",yanchor="bottom",y=1.01,x=0,font_size=12),
                          xaxis=dict(title="Projects",gridcolor="#F1F5F9",tickfont_size=11),
                          yaxis=dict(tickfont_size=11,automargin=True))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    with right:
        st.markdown("#### Year-over-year trend")
        yoy = (df_all[df_all["Year"].notna()]
               .groupby(["Year","Category"]).size().reset_index(name="Count"))
        yoy_t = df_all[df_all["Year"].notna()].groupby("Year").size().reset_index(name="Total")
        yoy = yoy.merge(yoy_t,on="Year")
        yoy["Pct"] = (yoy["Count"]/yoy["Total"]*100).round(1)
        fig2 = px.bar(yoy,x="Year",y="Pct",color="Category",
                      color_discrete_map=COLORS,barmode="stack",text="Pct",
                      category_orders={"Category":CAT_ORDER})
        fig2.update_traces(texttemplate="%{text:.0f}%",textposition="inside")
        fig2.update_layout(**CHART_CFG,height=240,
                           legend=dict(orientation="h",yanchor="bottom",y=1.01,font_size=11),
                           xaxis=dict(type="category",tickfont_size=11,title=""),
                           yaxis=dict(title="% of projects",range=[0,102],
                                      gridcolor="#F1F5F9",tickfont_size=11))
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar":False})

        st.markdown("#### SMS paid vs unpaid")
        sms_x = df.groupby(["Category","Paid SM Bool"]).size().reset_index(name="Count")
        sms_x["Status"] = sms_x["Paid SM Bool"].map({True:"SMS Paid ✓",False:"SMS Unpaid"})
        fig3 = px.bar(sms_x,x="Category",y="Count",color="Status",
                      color_discrete_map={"SMS Paid ✓":"#7C3AED","SMS Unpaid":"#C4B5FD"},
                      barmode="group",category_orders={"Category":CAT_ORDER})
        fig3.update_layout(**CHART_CFG,height=220,
                           legend=dict(orientation="h",yanchor="bottom",y=1.01,font_size=11),
                           xaxis=dict(tickfont_size=11,title=""),
                           yaxis=dict(title="Projects",gridcolor="#F1F5F9",tickfont_size=11))
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar":False})

    st.markdown("---")
    st.markdown("#### Theoretical vs. Real Margin · coloured by SMS paid status")
    ds = df[df["Theoretical Margin"].notna()&df["Real Margin"].notna()].copy()
    ds["SMS"] = ds["Paid SM Bool"].map({True:"SMS Paid ✓",False:"SMS Unpaid"})
    clip = ds[["Theoretical Margin","Real Margin"]].abs().quantile(0.98).max()
    ds = ds[(ds["Theoretical Margin"].abs()<=clip)&(ds["Real Margin"].abs()<=clip)]
    fig4 = px.scatter(ds,x="Theoretical Margin",y="Real Margin",color="SMS",
                      symbol="Category",
                      symbol_map={"Above 5%":"circle","Within Range (±5%)":"square","Less than 5%":"x"},
                      color_discrete_map={"SMS Paid ✓":"#7C3AED","SMS Unpaid":"#94A3B8"},
                      opacity=0.7,
                      hover_data={"Analytic Account":True,"Customer":True,"Salesperson":True},
                      labels={"Theoretical Margin":"Theoretical (€)","Real Margin":"Real (€)"})
    mn = min(ds["Theoretical Margin"].min(),ds["Real Margin"].min())
    mx = max(ds["Theoretical Margin"].max(),ds["Real Margin"].max())
    fig4.add_trace(go.Scatter(x=[mn,mx],y=[mn,mx],mode="lines",name="Perfect accuracy",
                              line=dict(color="#94A3B8",dash="dash",width=1.5),hoverinfo="skip"))
    fig4.add_trace(go.Scatter(x=[mn,mx],y=[mn*1.05,mx*1.05],mode="lines",name="+5% band",
                              line=dict(color="#86EFAC",dash="dot",width=1),hoverinfo="skip"))
    fig4.add_trace(go.Scatter(x=[mn,mx],y=[mn*0.95,mx*0.95],mode="lines",name="-5% band",
                              line=dict(color="#FCA5A5",dash="dot",width=1),hoverinfo="skip"))
    fig4.update_layout(**CHART_CFG,height=440,
                       legend=dict(orientation="h",yanchor="bottom",y=1.01,font_size=12),
                       xaxis=dict(gridcolor="#F1F5F9",tickfont_size=11),
                       yaxis=dict(gridcolor="#F1F5F9",tickfont_size=11))
    st.plotly_chart(fig4, use_container_width=True)


# ══ TAB 2 ══
with tab2:
    all_sp = sorted(df["Salesperson"].unique())
    c1,c2 = st.columns([2,1])
    with c1:
        sel_sp = st.selectbox("👤 Select Salesperson", all_sp)
    with c2:
        yr_opts = ["All Years"]+[str(y) for y in sorted(sel_years,reverse=True)]
        sp_yr = st.selectbox("🗓 Focus Year", yr_opts)

    df_sp = df[df["Salesperson"]==sel_sp].copy()
    if sp_yr != "All Years":
        df_sp = df_sp[df_sp["Year"]==int(sp_yr)]

    if df_sp.empty:
        st.info("No projects found for this selection.")
        st.stop()

    sp_above  = (df_sp["Category"]=="Above 5%").sum()
    sp_within = (df_sp["Category"]=="Within Range (±5%)").sum()
    sp_below  = (df_sp["Category"]=="Less than 5%").sum()
    sp_total  = len(df_sp)
    sp_theo   = df_sp["Theoretical Margin"].sum()
    sp_real   = df_sp["Real Margin"].sum()
    sp_var    = ((sp_real-sp_theo)/abs(sp_theo)*100) if abs(sp_theo)>1e-9 else 0
    sp_sms    = df_sp["Paid SM Bool"].sum()

    # Salesperson header
    st.markdown(f"## 👤 {sel_sp}")
    st.markdown(f"**{sp_total} projects** · Year: {sp_yr}")
    st.markdown("---")

    # KPI row
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("💰 Theoretical",  fmt(sp_theo))
    m2.metric("📈 Real Margin",  fmt(sp_real))
    m3.metric("📊 Variance",     f"{'+' if sp_var>=0 else ''}{sp_var:.1f}%")
    m4.metric("💬 SMS Paid",     f"{sp_sms} projects")

    st.markdown("<br>", unsafe_allow_html=True)

    # Category cards
    ca,cb,cc = st.columns(3)
    with ca:
        st.markdown(f'''<div style="background:#D1FAE5;border-radius:16px;padding:24px 28px;border-left:6px solid #16A34A;box-shadow:0 2px 8px rgba(0,0,0,0.07)">
        <p style="font-size:12px;font-weight:700;color:#15803D;text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px 0">🟢 ABOVE 5%</p>
        <p style="font-size:44px;font-weight:800;color:#14532D;margin:0;line-height:1">{sp_above:,}</p>
        <p style="font-size:14px;color:#166534;margin:8px 0 0 0;font-weight:500">{sp_above/sp_total*100:.1f}% of total</p></div>''', unsafe_allow_html=True)
    with cb:
        st.markdown(f'''<div style="background:#DBEAFE;border-radius:16px;padding:24px 28px;border-left:6px solid #2563EB;box-shadow:0 2px 8px rgba(0,0,0,0.07)">
        <p style="font-size:12px;font-weight:700;color:#1D4ED8;text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px 0">🔵 WITHIN ±5%</p>
        <p style="font-size:44px;font-weight:800;color:#1E3A8A;margin:0;line-height:1">{sp_within:,}</p>
        <p style="font-size:14px;color:#1E40AF;margin:8px 0 0 0;font-weight:500">{sp_within/sp_total*100:.1f}% of total</p></div>''', unsafe_allow_html=True)
    with cc:
        st.markdown(f'''<div style="background:#FEE2E2;border-radius:16px;padding:24px 28px;border-left:6px solid #DC2626;box-shadow:0 2px 8px rgba(0,0,0,0.07)">
        <p style="font-size:12px;font-weight:700;color:#DC2626;text-transform:uppercase;letter-spacing:.08em;margin:0 0 8px 0">🔴 LESS THAN 5%</p>
        <p style="font-size:44px;font-weight:800;color:#7F1D1D;margin:0;line-height:1">{sp_below:,}</p>
        <p style="font-size:14px;color:#991B1B;margin:8px 0 0 0;font-weight:500">{sp_below/sp_total*100:.1f}% of total</p></div>''', unsafe_allow_html=True)

    st.markdown("---")

    # Mini charts
    ch1,ch2 = st.columns(2,gap="large")
    with ch1:
        st.markdown("#### Category breakdown")
        fig_p = go.Figure(go.Pie(
            labels=CAT_ORDER, values=[sp_above,sp_within,sp_below], hole=0.6,
            marker_colors=[COLORS[c] for c in CAT_ORDER],
            textinfo="label+percent", textfont_size=12))
        fig_p.add_annotation(text=f"<b>{sp_total}</b><br>projects",
                             x=0.5,y=0.5,font_size=15,showarrow=False,font_color="#1A1F36")
        fig_p.update_layout(**CHART_CFG,height=260,showlegend=False)
        st.plotly_chart(fig_p, use_container_width=True, config={"displayModeBar":False})

    with ch2:
        st.markdown("#### Year-by-year trend")
        sp_yoy = (df_all[df_all["Salesperson"]==sel_sp]
                  .groupby(["Year","Category"]).size().reset_index(name="Count"))
        fig_y = px.bar(sp_yoy,x="Year",y="Count",color="Category",
                       color_discrete_map=COLORS,barmode="stack",
                       category_orders={"Category":CAT_ORDER})
        fig_y.update_layout(**CHART_CFG,height=260,
                            legend=dict(orientation="h",yanchor="bottom",y=1.01,font_size=11),
                            xaxis=dict(type="category",tickfont_size=11,title=""),
                            yaxis=dict(title="Projects",gridcolor="#F1F5F9",tickfont_size=11))
        st.plotly_chart(fig_y, use_container_width=True, config={"displayModeBar":False})

    st.markdown("---")

    # ── YEAR-WISE ANALYSIS ──
    st.markdown(f"#### 📅 Year-by-Year Performance Summary · {sel_sp}")
    sp_all_years = df_all[df_all["Salesperson"]==sel_sp].copy()
    years_list = sorted(sp_all_years["Year"].dropna().unique().tolist())
    year_rows = []
    for yr in years_list:
        ydf = sp_all_years[sp_all_years["Year"]==yr]
        n   = len(ydf)
        ab  = (ydf["Category"]=="Above 5%").sum()
        wi  = (ydf["Category"]=="Within Range (±5%)").sum()
        be  = (ydf["Category"]=="Less than 5%").sum()
        theo= ydf["Theoretical Margin"].sum()
        real= ydf["Real Margin"].sum()
        var = ((real-theo)/abs(theo)*100) if abs(theo)>1e-9 else 0
        sms = ydf["Paid SM Bool"].sum()
        year_rows.append({
            "Year": int(yr),
            "Total Projects": n,
            "🟢 Above 5%": ab,
            "🔵 Within ±5%": wi,
            "🔴 Below -5%": be,
            "Theoretical": fmt(theo),
            "Real Margin": fmt(real),
            "Variance %": f"{'+' if var>=0 else ''}{var:.1f}%",
            "💬 SMS Paid": sms,
        })
    yr_df = pd.DataFrame(year_rows)

    def style_var(v):
        try:
            f = float(str(v).replace("+","").replace("%",""))
            if f > 5:  return "color:#15803D;font-weight:700;background:#F0FDF4"
            if f < -5: return "color:#DC2626;font-weight:700;background:#FFF1F2"
            return "color:#2563EB;font-weight:700;background:#EFF6FF"
        except: return ""

    yr_styled = yr_df.style.map(style_var, subset=["Variance %"])
    st.dataframe(yr_styled, use_container_width=True, hide_index=True,
                 height=min(400,len(yr_df)*38+50))

    # Theo vs Real bar chart by year
    st.markdown(f"#### 📊 Theoretical vs Real Margin by Year · {sel_sp}")
    sp_yr_agg = sp_all_years.groupby("Year").agg(
        Theoretical=("Theoretical Margin","sum"),
        Real=("Real Margin","sum")).reset_index()
    sp_yr_agg["Year"] = sp_yr_agg["Year"].astype(str)
    fig_cmp = go.Figure()
    fig_cmp.add_trace(go.Bar(name="Theoretical",x=sp_yr_agg["Year"],y=sp_yr_agg["Theoretical"],
                             marker_color="#93C5FD",text=sp_yr_agg["Theoretical"].apply(fmt),
                             textposition="outside"))
    fig_cmp.add_trace(go.Bar(name="Real Margin",x=sp_yr_agg["Year"],y=sp_yr_agg["Real"],
                             marker_color="#34D399",text=sp_yr_agg["Real"].apply(fmt),
                             textposition="outside"))
    fig_cmp.update_layout(**CHART_CFG,barmode="group",height=300,
                          legend=dict(orientation="h",yanchor="bottom",y=1.01,font_size=12),
                          xaxis=dict(title="Year",tickfont_size=12),
                          yaxis=dict(title="Margin (€)",gridcolor="#F1F5F9",tickfont_size=11))
    st.plotly_chart(fig_cmp, use_container_width=True, config={"displayModeBar":False})

    st.markdown("---")

    # Project detail table
    st.markdown(f"#### 📋 Project Detail Table · {sel_sp}")
    search = st.text_input("🔍 Search projects",
                           placeholder="Filter by account or customer...",
                           label_visibility="collapsed")
    tbl = df_sp[["Analytic Account","Customer","State","Payment Status","Year",
                 "Theoretical Margin","Real Margin","Deviation",
                 "Category","Paid SM Bool","Paid SM Amount"]].copy()
    tbl["Deviation %"] = (tbl["Deviation"]*100).round(2)
    tbl["SMS Paid"] = tbl["Paid SM Bool"].map({True:"✅ Paid",False:"—"})
    tbl.drop(columns=["Deviation","Paid SM Bool"],inplace=True)
    tbl.rename(columns={"Theoretical Margin":"Theoretical (€)",
                        "Real Margin":"Real (€)",
                        "Paid SM Amount":"SMS Amount (€)"},inplace=True)

    # Color the Payment Status column
    def style_payment(v):
        if "Fully Paid"       in str(v): return "color:#15803D;font-weight:600"
        if "Invoiced"         in str(v): return "color:#D97706;font-weight:600"
        if "Not Yet Invoiced" in str(v): return "color:#DC2626;font-weight:600"
        return ""

    def style_state(v):
        if v == "Validated": return "color:#15803D;font-weight:600"
        if v == "Open":      return "color:#2563EB;font-weight:600"
        if v == "Submitted": return "color:#D97706;font-weight:600"
        if v == "Cancelled": return "color:#DC2626;font-weight:600"
        return ""
    if search:
        mask = (tbl["Analytic Account"].str.contains(search,case=False,na=False)|
                tbl["Customer"].str.contains(search,case=False,na=False))
        tbl = tbl[mask]

    def row_style(row):
        c = row.get("Category","")
        bg = "#F0FDF4" if "Above" in c else ("#FFF1F2" if "Less" in c else "#EFF6FF")
        fg = "#15803D" if "Above" in c else ("#DC2626" if "Less" in c else "#1D4ED8")
        return [f"background:{bg};color:{fg};font-weight:600" if col=="Category"
                else f"background:{bg}" for col in row.index]

    def dev_style(v):
        try:
            f = float(v)
            if f > 5:  return "color:#15803D;font-weight:700"
            if f < -5: return "color:#DC2626;font-weight:700"
            return "color:#2563EB;font-weight:700"
        except: return ""

    styled = (tbl.style.apply(row_style,axis=1)
              .map(dev_style,     subset=["Deviation %"])
              .map(style_payment, subset=["Payment Status"])
              .map(style_state,   subset=["State"])
              .format({"Theoretical (€)":"€{:,.2f}","Real (€)":"€{:,.2f}",
                       "SMS Amount (€)":"€{:,.2f}","Deviation %":"{:+.2f}%"},na_rep="—"))
    st.dataframe(styled, use_container_width=True,
                 height=min(560,max(280,len(tbl)*36+50)))
    st.caption(f"{len(tbl):,} projects · 🟢 Above 5%  🔵 Within ±5%  🔴 Less than 5%")
