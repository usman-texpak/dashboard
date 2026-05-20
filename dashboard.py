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
.stApp{background:#F4F6FB;}
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
div[data-testid="stMetric"]{background:white;border-radius:14px;padding:20px;border:1px solid #E8EDF5;}
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
    st.markdown("## 📊 Dashboard Controls")
    st.markdown("---")
    years_avail = sorted(df_all["Year"].dropna().unique().tolist(), reverse=True)
    sel_years = st.multiselect("📅 Year(s)", years_avail,
                               default=years_avail[:2] if len(years_avail)>=2 else years_avail)
    st.markdown("---")
    only_sms = st.toggle("💬 SMS Paid Only", value=False)
    st.markdown("---")
    sel_cats = st.multiselect("🎯 Categories", CAT_ORDER, default=CAT_ORDER)
    st.markdown("---")
    st.caption("Deviation = (Real − Theo) / |Theo| · 5% threshold per project")

df = df_all.copy()
if sel_years: df = df[df["Year"].isin(sel_years)]
if only_sms:  df = df[df["Paid SM Bool"]]
if sel_cats:  df = df[df["Category"].isin(sel_cats)]

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
        st.success(f"### 🟢 Above 5%\n# {above_n:,}\n**{pct_a:.1f}% of total projects**")
    with cb:
        st.info(f"### 🔵 Within ±5%\n# {within_n:,}\n**{pct_b:.1f}% of total projects**")
    with cc:
        st.error(f"### 🔴 Less than 5%\n# {below_n:,}\n**{pct_c:.1f}% of total projects**")

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
        st.success(f"### 🟢 Above 5%\n# {sp_above:,}\n**{sp_above/sp_total*100:.1f}% of total**")
    with cb:
        st.info(f"### 🔵 Within ±5%\n# {sp_within:,}\n**{sp_within/sp_total*100:.1f}% of total**")
    with cc:
        st.error(f"### 🔴 Less than 5%\n# {sp_below:,}\n**{sp_below/sp_total*100:.1f}% of total**")

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
    tbl = df_sp[["Analytic Account","Customer","State","Year",
                 "Theoretical Margin","Real Margin","Deviation",
                 "Category","Paid SM Bool","Paid SM Amount"]].copy()
    tbl["Deviation %"] = (tbl["Deviation"]*100).round(2)
    tbl["SMS Paid"] = tbl["Paid SM Bool"].map({True:"✅ Paid",False:"—"})
    tbl.drop(columns=["Deviation","Paid SM Bool"],inplace=True)
    tbl.rename(columns={"Theoretical Margin":"Theoretical (€)",
                        "Real Margin":"Real (€)",
                        "Paid SM Amount":"SMS Amount (€)"},inplace=True)
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
              .map(dev_style,subset=["Deviation %"])
              .format({"Theoretical (€)":"€{:,.2f}","Real (€)":"€{:,.2f}",
                       "SMS Amount (€)":"€{:,.2f}","Deviation %":"{:+.2f}%"},na_rep="—"))
    st.dataframe(styled, use_container_width=True,
                 height=min(560,max(280,len(tbl)*36+50)))
    st.caption(f"{len(tbl):,} projects · 🟢 Above 5%  🔵 Within ±5%  🔴 Less than 5%")
