import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re, requests
from io import BytesIO

st.set_page_config(page_title="SS Sale & Stock Analyzer", layout="wide", page_icon="📊")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Plus+Jakarta+Sans:wght@600;700;800&display=swap');
*, *::before, *::after { font-family: 'Inter', sans-serif !important; box-sizing: border-box; }
.stApp { background: #f4f0ff !important; }
/* ══ FILE UPLOADER ══ */
[data-testid="stFileUploader"] {
    background: #ffffff !important;
    border: none !important;
    border-radius: 14px !important;
    padding: 0 !important;
}
[data-testid="stFileUploaderDropzone"] {
    background: #ffffff !important;
    border: 2px dashed #c084fc !important;
    border-radius: 14px !important;
    text-align: center !important;
    padding: 16px !important;
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    min-height: 0 !important;
    max-height: 80px !important;
}
[data-testid="stFileUploaderDropzone"] svg { fill: #9c27b0 !important; }
[data-testid="stFileUploaderDropzone"] > div {
    display: flex !important;
    flex-direction: column !important;
    align-items: center !important;
    justify-content: center !important;
    width: 100% !important;
}
[data-testid="stFileUploaderDropzone"] button {
    visibility: hidden !important;
    height: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
}
[data-testid="stFileUploadDeleteBtn"] {
    visibility: visible !important;
    display: flex !important;
    align-items: center !important;
}
[data-testid="stFileUploadDeleteBtn"] button {
    visibility: visible !important;
    display: inline-flex !important;
    align-items: center !important;
    height: 28px !important;
    background: #fee2e2 !important;
    border: 1.5px solid #fca5a5 !important;
    border-radius: 6px !important;
    padding: 0 12px !important;
    cursor: pointer !important;
    color: #dc2626 !important;
    font-size: 12px !important;
    font-weight: 700 !important;
}
[data-testid="stFileUploadDeleteBtn"] button:hover {
    background: #fecaca !important;
}
[data-testid="stFileUploadDeleteBtn"] svg { fill: #dc2626 !important; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 0.8rem !important; padding-bottom: 1rem !important; }

.hero {
    padding: 0.55rem 1.4rem; display: flex; align-items: center; gap: 1rem;
    background: linear-gradient(90deg, #3a0068 0%, #6a1b9a 55%, #9c27b0 100%);
    margin-bottom: 1rem; border-radius: 12px;
    box-shadow: 0 3px 14px rgba(106,27,154,0.3);
}
.hero-badge {
    background: rgba(255,255,255,0.18); border: 1.5px solid rgba(255,255,255,0.35);
    color: #ffffff; font-size:.56rem; font-weight:700; letter-spacing:2px;
    text-transform:uppercase; padding:4px 11px; border-radius:20px; white-space:nowrap;
}
.hero-title { font-family:'Plus Jakarta Sans',sans-serif !important; font-size:1.05rem; font-weight:800; color:#ffffff; }
.hero-sub-line { font-family:'Plus Jakarta Sans',sans-serif !important; font-size:.8rem; font-weight:600; color:#e8c8ff; }
.hero-sub { color:rgba(255,255,255,0.52); font-size:.65rem; font-weight:400; }

.kpi-card {
    background: linear-gradient(135deg, #6a1b9a 0%, #9c27b0 100%);
    border-radius: 16px; padding: 1.1rem 1.3rem;
    box-shadow: 0 4px 18px rgba(106,27,154,0.35);
}
.kpi-label { font-size:.58rem; font-weight:700; letter-spacing:2.5px; text-transform:uppercase; color:rgba(255,255,255,0.75); margin-bottom:.4rem; }
.kpi-value { font-family:'Plus Jakarta Sans',sans-serif !important; font-size:1.3rem; font-weight:800; color:#ffffff; line-height:1.1; }
.kpi-sub { font-size:.72rem; color:rgba(255,255,255,0.7); margin-top:.25rem; }

.section-title {
    font-size:.63rem; font-weight:700; letter-spacing:2.5px; text-transform:uppercase;
    color:#6a1b9a; padding:.4rem 0; margin-bottom:.7rem;
    border-bottom: 2px solid #ddd6fe;
}

p { color:#1a0030 !important; font-size:.9rem !important; }
label { color:#3d0066 !important; font-weight:600 !important; }
[data-testid="stWidgetLabel"] p { color:#6a1b9a !important; font-weight:600 !important; }

.stSelectbox > div > div, .stMultiSelect > div > div {
    background:#ffffff !important; border:1.5px solid #c084fc !important;
    border-radius:10px !important; color:#1a0030 !important;
}
[data-baseweb="popover"] { background:#fff !important; }
[data-baseweb="popover"] * { color:#1a0030 !important; background:#fff !important; }
[data-baseweb="tag"] { background:#ede9fe !important; }
[data-baseweb="tag"] span { color:#4c1d95 !important; font-weight:600 !important; }

.stButton > button {
    background: linear-gradient(135deg,#6a1b9a,#9c27b0) !important;
    color:#ffffff !important; border:none !important; border-radius:12px !important;
    font-weight:700 !important; padding:.7rem 2rem !important;
    box-shadow:0 4px 14px rgba(106,27,154,0.38) !important;
}
.stButton > button p { color:#ffffff !important; font-weight:700 !important; }
.stDownloadButton > button {
    background:#fff !important; color:#6a1b9a !important;
    border:2px solid #6a1b9a !important; border-radius:10px !important; font-weight:700 !important;
}


.stTabs [data-baseweb="tab-list"] {
    background:#fff !important; border-radius:12px !important; padding:4px !important;
    border:1.5px solid #ddd6fe !important;
}
.stTabs [data-baseweb="tab"] { color:#1a0030 !important; border-radius:8px !important; font-size:.84rem !important; font-weight:600 !important; }
.stTabs [aria-selected="true"] { background:linear-gradient(135deg,#6a1b9a,#9c27b0) !important; color:#ffffff !important; }
.stTabs [aria-selected="true"] * { color:#ffffff !important; }
.stSuccess * { color:#166534 !important; font-weight:600 !important; }
div[data-testid="stDataFrame"] * { color:#1a0030 !important; font-size:.84rem !important; }
</style>
""", unsafe_allow_html=True)

# ══ CONSTANTS ══
MONTHS_ORDER = ["Apr'25","May'25","June'25","Jul'25","Aug'25","Sep'25","Oct'25","Nov'25","Dec'25","Jan'26","Feb'26"]
MONTH_SHORT  = ["Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb"]
BRANDS_MAIN  = ['ADIDAS','ASICS','CROCS','DCYPHR','LEVIS','NIKE','OTHERS','PUMA','SKECHERS']
DIVISIONS    = ['FOOTWEAR','APPAREL','ACCESSORIES']
CATEGORIES   = ['LIFESTYLE','RUNNING/TRAINING','SOCCER & SPORTS','WALKING','MOTORSPORT','ACTIVE WEAR']

CAT_COLORS = ['#7b1fa2','#e91e63','#ff6f00','#1565c0','#2e7d32','#00838f',
              '#f57f17','#6a1b9a','#c62828','#00695c','#4527a0','#ad1457']
BLUE_SEQ   = [[0,'#f3e5f5'],[0.4,'#9c27b0'],[1,'#6a1b9a']]

# ══ HELPERS ══
def fmt_lac(v):
    if pd.isna(v) or v == 0: return "—"
    v = float(v)
    neg = v < 0
    v = abs(v)
    if v >= 100:
        s = str(int(round(v)))
        if len(s) <= 3: r = s
        else:
            last3 = s[-3:]; rest = s[:-3]; grp = []
            while len(rest) > 2: grp.append(rest[-2:]); rest = rest[:-2]
            if rest: grp.append(rest)
            grp.reverse()
            r = ','.join(grp) + ',' + last3
        return ('-' if neg else '') + r + ' L'
    return ('-' if neg else '') + f"{v:.2f} L"

def pct(v, dec=2):
    if pd.isna(v) or v == 0: return "—"
    return f"{float(v)*100:.{dec}f}%"

def cl(height=380, title="", xangle=0, show_legend=True, margin=None):
    """chart_layout — clean, no conflicts"""
    m = margin or dict(l=10, r=10, t=55, b=40)
    return dict(
        paper_bgcolor="rgba(255,255,255,1)",
        plot_bgcolor="rgba(245,240,255,0.5)",
        font=dict(color="#1a0030", family="Inter", size=12),
        margin=m, height=height,
        title=dict(text=f"<b>{title}</b>", font=dict(color="#1a0030", size=14, family="Plus Jakarta Sans")),
        legend=dict(font=dict(color="#1a0030", size=11), bgcolor="rgba(255,255,255,0.97)",
                    bordercolor="#ddd6fe", borderwidth=1.5, visible=show_legend),
        xaxis=dict(gridcolor="#ede9fe", tickfont=dict(color="#1a0030", size=11),
                   linecolor="#ddd6fe", tickangle=xangle, showgrid=True),
        yaxis=dict(gridcolor="#ede9fe", tickfont=dict(color="#1a0030", size=11),
                   linecolor="#ddd6fe", showgrid=True),
    )

def kpi_card(col, label, value, sub, icon):
    with col:
        st.markdown(f"""<div class="kpi-card">
          <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:.3rem">
            <div class="kpi-label">{label}</div><span style="font-size:1.1rem">{icon}</span>
          </div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

def sec(title):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)

def normalize_brand(b):
    if pd.isna(b): return 'OTHERS'
    b = str(b).strip().upper()
    known = ['ADIDAS','ASICS','CROCS','DCYPHR','LEVIS','NIKE','PUMA','SKECHERS']
    return b if b in known else 'OTHERS'

def is_fw(s):
    try: v = float(str(s).replace('½','.5')); return 5 <= v <= 13
    except: return False

def is_ap(s):
    return str(s).upper().strip() in ['XS','S','M','L','XL','XXL','XXXL','2XL','3XL','3X','2X']

AP_ORDER = ['XS','S','M','L','XL','XXL','XXXL','2XL','3XL','3X','2X']
AP_MAP   = {s:i for i,s in enumerate(AP_ORDER)}
AP_REV   = {i:s for i,s in enumerate(AP_ORDER)}

# ══ PROCESS ══
@st.cache_data(show_spinner=False)
def process(file):
    stock = pd.read_excel(file, sheet_name="STOCK REPORT", header=1)
    sale  = pd.read_excel(file, sheet_name="SALE REPORT",  header=1)
    stock.columns = [str(c).strip() for c in stock.columns]
    sale.columns  = [str(c).strip() for c in sale.columns]

    stock = stock.rename(columns={
        'Divison  Desc':'Division','Category Desc':'Category',
        'GENDER.':'Gender','Brand.':'Brand','Clsoing Value':'StockValue'
    })
    sale = sale.rename(columns={
        'Item No/ Article Code':'ItemID','Divison  Desc':'Division',
        'Category Code':'Category','GENDER.':'Gender','Brand.':'Brand',
        'Net Sale Tax Incl':'NetSale','Mrp Value':'MrpValue'
    })

    for c in ['MRP','Closing Qty','StockValue','GIT']:
        if c in stock.columns: stock[c] = pd.to_numeric(stock[c], errors='coerce').fillna(0)
    for c in ['Sale Qty','NetSale','MrpValue','Disc value']:
        if c in sale.columns:  sale[c]  = pd.to_numeric(sale[c],  errors='coerce').fillna(0)

    # Values already in Lacs — keep as is
    stock['Brand']  = stock['Brand'].apply(normalize_brand)
    sale['Brand']   = sale['Brand'].apply(normalize_brand)
    stock['Gender'] = stock['Gender'].str.upper().str.strip()
    sale['Gender']  = sale['Gender'].str.upper().str.strip()

    all_stores  = sorted(set(sale['Store Name'].dropna()) | set(stock['Store Name'].dropna()))
    grand_sale  = sale['NetSale'].sum()
    grand_stock = stock['StockValue'].sum()
    grand_qty   = stock['Closing Qty'].sum()
    return sale, stock, all_stores, grand_sale, grand_stock, grand_qty

# ══ CACHED HELPERS ══
@st.cache_data(show_spinner=False)
def get_store_sale(sale_df, store):
    return sale_df[sale_df['Store Name']==store].copy()

@st.cache_data(show_spinner=False)
def get_store_stock(stock_df, store):
    return stock_df[stock_df['Store Name']==store].copy()

@st.cache_data(show_spinner=False)
def get_month_data(sale_df, store, month_order):
    ss = sale_df[sale_df['Store Name']==store]
    return ss[ss['Month']==month_order].copy()

# ══ SESSION ══
for k,v in {"ready":False,"data":None,"ai_text":None,"ai_loading":False}.items():
    if k not in st.session_state: st.session_state[k] = v

# ══ HERO ══
st.markdown("""
<div class="hero">
  <div class="hero-badge">SS Analyzer</div>
  <div style="width:1px;height:26px;background:rgba(255,255,255,.22)"></div>
  <div>
    <div style="display:flex;align-items:baseline;gap:.6rem">
      <div class="hero-title">SS Sale &amp; Stock Analyzer</div>
      <div style="color:rgba(255,255,255,.4)">→</div>
      <div class="hero-sub-line">Store · Brand · Gender · Category · Size · Cut Size</div>
    </div>
    <div class="hero-sub">Upload RAW_DATA_REPORTS_INSIGHT.xlsx &nbsp;·&nbsp; Auto Reports &nbsp;·&nbsp; Interactive Dashboard</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══ UPLOAD ══
u1,u2,u3 = st.columns([1,2,1])
with u2:
    uploaded = st.file_uploader(
        "📂 Upload RAW_DATA_REPORTS_INSIGHT.xlsx",
        type=["xlsx","xls"],
        label_visibility="visible"
    )
    if uploaded:
        if st.button("⚡  Generate Reports + Dashboard", use_container_width=True):
            with st.spinner("Processing..."):
                st.session_state.data    = process(uploaded)
                st.session_state.ready   = True
                st.session_state.ai_text = None
            st.success("✅ Done! Reports Ready.")

if not st.session_state.ready:
    st.markdown("""<div style="text-align:center;padding:5rem 0">
      <div style="font-size:3.5rem">📊</div>
      <div style="margin-top:1rem;font-size:1rem;color:#607d9b;font-weight:500">Upload file and click Generate</div>
    </div>""", unsafe_allow_html=True)
    st.stop()

sale, stock, all_stores, grand_sale, grand_stock, grand_qty = st.session_state.data

# ══ KPIs ══
top_store = sale.groupby('Store Name')['NetSale'].sum().idxmax()
top_brand = sale.groupby('Brand')['NetSale'].sum().idxmax()
top_store_val = sale.groupby('Store Name')['NetSale'].sum().max()
top_brand_val = sale.groupby('Brand')['NetSale'].sum().max()

k1,k2,k3,k4,k5 = st.columns(5)
kpi_card(k1,"Total Sale",    fmt_lac(grand_sale),    f"Apr'25–Feb'26 · {len(all_stores)} Stores","💰")
kpi_card(k2,"Closing Stock", fmt_lac(grand_stock),   f"Feb 2026 · {int(grand_qty):,} Pcs","📦")
kpi_card(k3,"Top Store",     top_store[:22],          fmt_lac(top_store_val),"🏆")
kpi_card(k4,"Top Brand",     top_brand,               fmt_lac(top_brand_val),"🏷️")
kpi_card(k5,"Brands",        str(len(sale['Brand'].unique())), "Active Brands","⭐")
st.markdown("<br>", unsafe_allow_html=True)

# ══ TABS ══
t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11 = st.tabs([
    "📈 Overview","🏪 Store-wise","🏷️ Brand-wise","👤 Gender-wise",
    "🗂️ Category","📐 Size & Cut Size","🔥 Heatmap","🔍 Store Deep Dive",
    "📊 Performance","📦 Inventory","🤖 AI Strategy"
])

# ══ TAB 1: OVERVIEW ══
with t1:
    sec("📈 Monthly Sale Trend")
    monthly = sale.groupby('Month')['NetSale'].sum().reindex(MONTHS_ORDER).fillna(0)
    bi = int(monthly.values.argmax()); wi = int(monthly.values.argmin())
    bcolors = ['#9c27b0' if i not in [bi,wi] else ('#16a34a' if i==bi else '#dc2626') for i in range(len(monthly))]
    fig = go.Figure(go.Bar(
        x=MONTH_SHORT, y=monthly.values,
        marker=dict(color=bcolors, line=dict(width=0)),
        text=[fmt_lac(v) for v in monthly.values],
        textposition='outside', textfont=dict(size=11,color='#1a0030'),
    ))
    fig.update_layout(**cl(380,"Monthly Net Sale — All Stores (Apr'25–Feb'26)",
        margin=dict(l=10,r=10,t=55,b=40)),
        bargap=0.3, yaxis_range=[0,monthly.max()*1.22],
        annotations=[
            dict(x=MONTH_SHORT[bi],y=monthly.values[bi]*1.15,text="🏆 Best",showarrow=False,font=dict(color='#16a34a',size=11)),
            dict(x=MONTH_SHORT[wi],y=monthly.values[wi]*1.15,text="⬇ Low",showarrow=False,font=dict(color='#dc2626',size=11)),
        ])
    st.plotly_chart(fig, use_container_width=True)

    avg_m = monthly.mean()
    feb_g = ((monthly.values[-1]-monthly.values[0])/monthly.values[0]*100) if monthly.values[0]>0 else 0
    i1,i2,i3,i4 = st.columns(4)
    for col,lbl,val,sub in [
        (i1,"📈 Best Month",  MONTH_SHORT[bi],         fmt_lac(monthly.values[bi])),
        (i2,"📉 Lowest Month",MONTH_SHORT[wi],         fmt_lac(monthly.values[wi])),
        (i3,"📊 Avg Monthly", fmt_lac(avg_m),          "Per month average"),
        (i4,"🚀 Apr→Feb",     f"{feb_g:+.1f}%",        "Growth trend"),
    ]:
        with col:
            st.markdown(f"""<div style="background:#fff;border-radius:12px;padding:.9rem 1.1rem;
                box-shadow:0 2px 10px rgba(106,27,154,.1);border-left:4px solid #9c27b0">
                <div style="font-size:.58rem;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#6a1b9a">{lbl}</div>
                <div style="font-size:1.25rem;font-weight:800;color:#1a0030;margin:.2rem 0">{val}</div>
                <div style="font-size:.72rem;color:#607d9b">{sub}</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    ca,cb = st.columns(2)
    with ca:
        sec("🏆 Top 10 Stores")
        top10 = sale.groupby('Store Name')['NetSale'].sum().nlargest(10).sort_values()
        fig2 = go.Figure(go.Bar(
            x=top10.values, y=top10.index.tolist(), orientation='h',
            marker=dict(color=top10.values, colorscale=BLUE_SEQ, line=dict(width=0)),
            text=[fmt_lac(v) for v in top10.values],
            textposition='outside', textfont=dict(size=11,color='#1a0030'),
        ))
        fig2.update_layout(**cl(420,"Top 10 Stores by Net Sale",
            margin=dict(l=10,r=160,t=55,b=40)),
            xaxis_range=[0,top10.max()*1.45])
        st.plotly_chart(fig2, use_container_width=True)

    with cb:
        sec("🏷️ Brand Mix")
        bs = sale.groupby('Brand')['NetSale'].sum().sort_values(ascending=False)
        fig3 = go.Figure(go.Pie(
            labels=bs.index.tolist(), values=bs.values.tolist(), hole=0.52,
            marker=dict(colors=CAT_COLORS[:len(bs)], line=dict(color='#fff',width=2)),
            textinfo='label+percent', textfont=dict(size=11,color='#1a0030'),
            insidetextfont=dict(size=10,color='#fff'),
        ))
        fig3.update_layout(**cl(420,"Brand-wise Sale Contribution",margin=dict(l=10,r=10,t=55,b=10)),
            annotations=[dict(text=f"<b>{fmt_lac(grand_sale)}</b>",x=0.5,y=0.5,
                              font=dict(size=12,color='#1a0030'),showarrow=False)])
        st.plotly_chart(fig3, use_container_width=True)

# ══ TAB 2: STORE-WISE ══
with t2:
    sec("🏪 Store-wise Monthly Sale")
    swc = sale.pivot_table(index='Store Name',columns='Month',values='NetSale',aggfunc='sum').reindex(columns=MONTHS_ORDER).fillna(0)
    swc['Total Sale']    = swc[MONTHS_ORDER].sum(axis=1)
    swc['Closing Stock'] = stock.groupby('Store Name')['StockValue'].sum()
    swc['Closing Qty']   = stock.groupby('Store Name')['Closing Qty'].sum()
    swc['Sale Cont.']    = swc['Total Sale'] / swc['Total Sale'].sum()

    top5s = swc['Total Sale'].nlargest(5).index.tolist()
    sel   = st.multiselect("Select Stores", swc.index.tolist(), default=top5s, key="swc_sel")
    if sel:
        fig_s = go.Figure()
        for i,sn in enumerate(sel):
            if sn in swc.index:
                fig_s.add_trace(go.Bar(x=MONTH_SHORT, y=swc.loc[sn,MONTHS_ORDER].values,
                    name=sn, marker_color=CAT_COLORS[i%len(CAT_COLORS)],
                    hovertemplate=f'<b>{sn}</b><br>%{{x}}: %{{y:.2f}} L<extra></extra>'))
        fig_s.update_layout(**cl(430,"Store-wise Monthly Sale",margin=dict(l=10,r=10,t=55,b=40)),
            barmode='group', bargap=0.12)
        st.plotly_chart(fig_s, use_container_width=True)

    sec("📋 SWC Table")
    disp = swc.copy()
    disp.index.name = "Store Name"
    for c in MONTHS_ORDER: disp[c] = disp[c].apply(lambda x: round(x,2) if x!=0 else "")
    disp['Total Sale']    = disp['Total Sale'].apply(lambda x: round(x,2) if x!=0 else "")
    disp['Closing Stock'] = disp['Closing Stock'].apply(lambda x: round(x,2) if pd.notna(x) and x!=0 else "")
    disp['Closing Qty']   = disp['Closing Qty'].apply(lambda x: int(x) if pd.notna(x) and x!=0 else "")
    disp['Sale Cont.']    = disp['Sale Cont.'].apply(lambda x: pct(x,2) if pd.notna(x) else "")
    st.dataframe(disp, use_container_width=True)

# ══ TAB 3: BRAND-WISE ══
with t3:
    sec("🏷️ Brand-wise Monthly Sale Trend")
    bwc = sale.pivot_table(index='Brand',columns='Month',values='NetSale',aggfunc='sum').reindex(columns=MONTHS_ORDER).fillna(0)
    bwc['Total Sale']    = bwc[MONTHS_ORDER].sum(axis=1)
    bwc['Closing Stock'] = stock.groupby('Brand')['StockValue'].sum()
    bwc['Closing Qty']   = stock.groupby('Brand')['Closing Qty'].sum()
    bwc['Sale Cont.']    = bwc['Total Sale'] / bwc['Total Sale'].sum()
    bwc = bwc.sort_values('Total Sale',ascending=False)

    fig_bl = go.Figure()
    for i,b in enumerate(bwc.index):
        fig_bl.add_trace(go.Scatter(
            x=MONTH_SHORT, y=bwc.loc[b,MONTHS_ORDER].values, name=b,
            mode='lines+markers', line=dict(color=CAT_COLORS[i%len(CAT_COLORS)],width=2.5),
            marker=dict(size=7),
            hovertemplate=f'<b>{b}</b><br>%{{x}}: %{{y:.2f}} L<extra></extra>'))
    fig_bl.update_layout(**cl(420,"Brand-wise Monthly Sale Trend",margin=dict(l=10,r=10,t=55,b=40)))
    st.plotly_chart(fig_bl, use_container_width=True)

    ba,bb = st.columns(2)
    with ba:
        sec("📊 Brand Sale vs Stock")
        brands = bwc.index.tolist()
        sv = bwc['Total Sale'].values.tolist()
        kv = [float(bwc.loc[b,'Closing Stock']) if pd.notna(bwc.loc[b,'Closing Stock']) else 0 for b in brands]
        fig_bvk = go.Figure()
        fig_bvk.add_trace(go.Bar(name='Sale',  x=brands, y=sv, marker_color='#7b1fa2',
            text=[fmt_lac(v) for v in sv], textposition='outside', textfont=dict(size=10,color='#4a0072')))
        fig_bvk.add_trace(go.Bar(name='Stock', x=brands, y=kv, marker_color='#ce93d8',
            text=[fmt_lac(v) for v in kv], textposition='outside', textfont=dict(size=10,color='#6b21a8')))
        fig_bvk.update_layout(**cl(380,"Brand: Sale vs Closing Stock",margin=dict(l=10,r=10,t=55,b=70)),
            barmode='group', bargap=0.2, xaxis_tickangle=-30)
        st.plotly_chart(fig_bvk, use_container_width=True)

    with bb:
        sec("📐 Brand Sell-Through Rate")
        st_r = []
        for b in brands:
            s = float(bwc.loc[b,'Total Sale'])
            k = float(bwc.loc[b,'Closing Stock']) if pd.notna(bwc.loc[b,'Closing Stock']) else 0
            t = s+k; st_r.append(round(s/t*100,1) if t>0 else 0)
        colors_st = ['#16a34a' if r>=60 else ('#ca8a04' if r>=30 else '#dc2626') for r in st_r]
        fig_bst = go.Figure(go.Bar(
            x=brands, y=st_r, marker=dict(color=colors_st,line=dict(width=0)),
            text=[f"{r:.1f}%" for r in st_r], textposition='outside', textfont=dict(size=11,color='#1a0030')))
        fig_bst.update_layout(**cl(380,"Brand Sell-Through Rate (%)",margin=dict(l=10,r=10,t=55,b=70)),
            bargap=0.3, yaxis_range=[0,115], xaxis_tickangle=-30,
            shapes=[
                dict(type='line',x0=-0.5,x1=len(brands)-0.5,y0=60,y1=60,line=dict(color='#16a34a',width=2,dash='dash')),
                dict(type='line',x0=-0.5,x1=len(brands)-0.5,y0=30,y1=30,line=dict(color='#dc2626',width=2,dash='dash'))
            ])
        st.plotly_chart(fig_bst, use_container_width=True)

    sec("📋 Brand Summary Table")
    bd = bwc[['Total Sale','Closing Stock','Closing Qty','Sale Cont.']].copy()
    bd['Total Sale']    = bd['Total Sale'].apply(lambda x: round(x,2) if x!=0 else "")
    bd['Closing Stock'] = bd['Closing Stock'].apply(lambda x: round(x,2) if pd.notna(x) and x!=0 else "")
    bd['Closing Qty']   = bd['Closing Qty'].apply(lambda x: int(x) if pd.notna(x) and x!=0 else "")
    bd['Sale Cont.']    = bd['Sale Cont.'].apply(lambda x: pct(x,2) if pd.notna(x) else "")
    st.dataframe(bd, use_container_width=True)

# ══ TAB 4: GENDER-WISE ══
with t4:
    gdr_s = sale.groupby('Gender')['NetSale'].sum().sort_values(ascending=False)
    gdr_k = stock.groupby('Gender')['StockValue'].sum()

    ga,gb = st.columns(2)
    with ga:
        sec("👤 Gender Sale Distribution")
        fig_gp = go.Figure(go.Pie(
            labels=gdr_s.index.tolist(), values=gdr_s.values.tolist(), hole=0.5,
            marker=dict(colors=['#7b1fa2','#e91e63','#1565c0','#ff6f00'],line=dict(color='#fff',width=2)),
            textinfo='label+percent', textfont=dict(size=12,color='#1a0030'),
            insidetextfont=dict(size=10,color='#fff')))
        fig_gp.update_layout(**cl(380,"Gender-wise Sale Distribution",margin=dict(l=10,r=10,t=55,b=10)))
        st.plotly_chart(fig_gp, use_container_width=True)

    with gb:
        sec("📊 Gender: Sale vs Stock")
        gens = gdr_s.index.tolist()
        sv_g = [float(gdr_s.get(g,0)) for g in gens]
        kv_g = [float(gdr_k.get(g,0)) for g in gens]
        fig_gb = go.Figure()
        fig_gb.add_trace(go.Bar(name='Sale', x=gens, y=sv_g, marker_color='#7b1fa2',
            text=[fmt_lac(v) for v in sv_g], textposition='outside'))
        fig_gb.add_trace(go.Bar(name='Stock',x=gens, y=kv_g, marker_color='#ce93d8',
            text=[fmt_lac(v) for v in kv_g], textposition='outside'))
        fig_gb.update_layout(**cl(380,"Gender: Sale vs Stock",margin=dict(l=10,r=10,t=55,b=40)),
            barmode='group', bargap=0.25)
        st.plotly_chart(fig_gb, use_container_width=True)

    sec("📅 Gender Monthly Trend")
    gm = sale.pivot_table(index='Gender',columns='Month',values='NetSale',aggfunc='sum').reindex(columns=MONTHS_ORDER).fillna(0)
    fig_gm = go.Figure()
    for i,g in enumerate(gm.index):
        fig_gm.add_trace(go.Scatter(x=MONTH_SHORT, y=gm.loc[g].values, name=g,
            mode='lines+markers', line=dict(color=CAT_COLORS[i%4],width=2.5), marker=dict(size=7)))
    fig_gm.update_layout(**cl(360,"Gender Monthly Trend",margin=dict(l=10,r=10,t=55,b=40)))
    st.plotly_chart(fig_gm, use_container_width=True)

    sec("🏷️ Gender × Brand Matrix")
    gbm = sale.pivot_table(index='Gender',columns='Brand',values='NetSale',aggfunc='sum').fillna(0)
    gbm['TOTAL'] = gbm.sum(axis=1)
    st.dataframe(gbm.map(lambda x: round(x,2) if x!=0 else ""), use_container_width=True)

# ══ TAB 5: CATEGORY ══
with t5:
    div_s = sale.groupby('Division')['NetSale'].sum().sort_values(ascending=False)
    div_k = stock.groupby('Division')['StockValue'].sum()
    da,db = st.columns(2)
    with da:
        sec("🗂️ Division-wise Sale")
        fig_dv = go.Figure(go.Bar(
            x=div_s.index.tolist(), y=div_s.values.tolist(),
            marker=dict(color=CAT_COLORS[:len(div_s)],line=dict(width=0)),
            text=[fmt_lac(v) for v in div_s.values], textposition='outside', textfont=dict(size=12,color='#1a0030')))
        fig_dv.update_layout(**cl(340,"Division-wise Sale",margin=dict(l=10,r=10,t=55,b=40)),
            bargap=0.4, yaxis_range=[0,div_s.max()*1.22])
        st.plotly_chart(fig_dv, use_container_width=True)

    with db:
        sec("📊 Division: Sale vs Stock")
        dvs = div_s.index.tolist()
        fig_dvk = go.Figure()
        fig_dvk.add_trace(go.Bar(name='Sale', x=dvs,y=[float(div_s.get(d,0)) for d in dvs],marker_color='#7b1fa2',
            text=[fmt_lac(float(div_s.get(d,0))) for d in dvs],textposition='outside'))
        fig_dvk.add_trace(go.Bar(name='Stock',x=dvs,y=[float(div_k.get(d,0)) for d in dvs],marker_color='#ce93d8',
            text=[fmt_lac(float(div_k.get(d,0))) for d in dvs],textposition='outside'))
        fig_dvk.update_layout(**cl(340,"Division: Sale vs Stock",margin=dict(l=10,r=10,t=55,b=40)),
            barmode='group',bargap=0.3)
        st.plotly_chart(fig_dvk, use_container_width=True)

    cat_s = sale.groupby('Category')['NetSale'].sum().sort_values(ascending=False)
    cat_k = stock.groupby('Category')['StockValue'].sum()
    ca5,cb5 = st.columns(2)
    with ca5:
        sec("📦 Category-wise Sale")
        fig_cs = go.Figure(go.Bar(
            x=cat_s.index.tolist(), y=cat_s.values.tolist(),
            marker=dict(color=CAT_COLORS[:len(cat_s)],line=dict(width=0)),
            text=[fmt_lac(v) for v in cat_s.values], textposition='outside', textfont=dict(size=11,color='#1a0030')))
        fig_cs.update_layout(**cl(360,"Category-wise Sale",margin=dict(l=10,r=10,t=55,b=90)),
            bargap=0.3, yaxis_range=[0,cat_s.max()*1.22], xaxis_tickangle=-30)
        st.plotly_chart(fig_cs, use_container_width=True)

    with cb5:
        sec("📐 Category Sell-Through %")
        cats5 = cat_s.index.tolist()
        sv5 = [float(cat_s.get(c,0)) for c in cats5]
        kv5 = [float(cat_k.get(c,0)) for c in cats5]
        str5 = [round(sv5[i]/(sv5[i]+kv5[i])*100,1) if (sv5[i]+kv5[i])>0 else 0 for i in range(len(cats5))]
        fig_cst = go.Figure(go.Bar(
            x=cats5, y=str5, marker=dict(color=['#16a34a' if r>=60 else ('#ca8a04' if r>=30 else '#dc2626') for r in str5],line=dict(width=0)),
            text=[f"{r:.1f}%" for r in str5], textposition='outside', textfont=dict(size=11,color='#1a0030')))
        fig_cst.update_layout(**cl(360,"Category Sell-Through Rate (%)",margin=dict(l=10,r=10,t=55,b=90)),
            bargap=0.3, yaxis_range=[0,115], xaxis_tickangle=-30)
        st.plotly_chart(fig_cst, use_container_width=True)

    sec("🔍 Filter: Division → Category → Brand")
    f1,f2,f3 = st.columns(3)
    with f1: div_f = st.selectbox("Division",["All"]+DIVISIONS,key="cat_div")
    with f2: cat_f = st.selectbox("Category",["All"]+CATEGORIES,key="cat_cat")
    with f3: brd_f = st.selectbox("Brand",["All"]+sorted(sale['Brand'].unique()),key="cat_brd")
    sf = sale.copy()
    if div_f!="All": sf=sf[sf['Division']==div_f]
    if cat_f!="All": sf=sf[sf['Category']==cat_f]
    if brd_f!="All": sf=sf[sf['Brand']==brd_f]
    fm = sf.groupby('Month')['NetSale'].sum().reindex(MONTHS_ORDER).fillna(0)
    fig_ff = go.Figure(go.Bar(x=MONTH_SHORT,y=fm.values,marker=dict(color='#7b1fa2',line=dict(width=0)),
        text=[fmt_lac(v) for v in fm.values],textposition='outside',textfont=dict(size=11,color='#1a0030')))
    fig_ff.update_layout(**cl(300,f"Monthly Sale — {div_f}|{cat_f}|{brd_f}",margin=dict(l=10,r=10,t=55,b=40)),
        bargap=0.3, yaxis_range=[0,max(fm.max()*1.2,1)])
    st.plotly_chart(fig_ff, use_container_width=True)

    sec("📋 Store × Category Table")
    sc_tbl = sale.pivot_table(index='Store Name',columns='Category',values='NetSale',aggfunc='sum').fillna(0)
    sc_tbl['TOTAL'] = sc_tbl.sum(axis=1)
    sc_tbl = sc_tbl.sort_values('TOTAL',ascending=False)
    st.dataframe(sc_tbl.map(lambda x: round(x,2) if x!=0 else ""), use_container_width=True)

# ══ TAB 6: SIZE & CUT SIZE ══
with t6:

    # ── FULL SIZE vs CUT SIZE SECTION ──
    sec("✂️ Full Size & Cut Size SKU Analysis")

    # Filters
    fc1, fc2, fc3, fc4 = st.columns(4)
    with fc1:
        fs_store = st.selectbox("🏪 Filter by Store", ["All"] + sorted(stock["Store Name"].dropna().unique()), key="fs_store")
    with fc2:
        fs_brand = st.selectbox("🏷️ Filter by Brand", ["All"] + sorted(stock["Brand"].dropna().unique()), key="fs_brand")
    with fc3:
        fs_gender = st.selectbox("👤 Filter by Gender", ["All", "MEN", "WOMEN", "KIDS"], key="fs_gender")
    with fc4:
        fs_div = st.selectbox("📦 Filter by Division", ["All"] + DIVISIONS, key="fs_div")

    # Apply filters
    stk_fs = stock.copy()
    if fs_store  != "All": stk_fs = stk_fs[stk_fs["Store Name"] == fs_store]
    if fs_brand  != "All": stk_fs = stk_fs[stk_fs["Brand"]      == fs_brand]
    if fs_gender != "All": stk_fs = stk_fs[stk_fs["Gender"]     == fs_gender]
    if fs_div    != "All": stk_fs = stk_fs[stk_fs["Division"]   == fs_div]

    # Dynamic description based on filters
    if fs_div == "FOOTWEAR":
        if fs_gender == "WOMEN":
            size_info = "Women Footwear sizes checked: <b>4, 5, 6, 7</b>"
        else:
            size_info = "Men Footwear sizes checked: <b>7, 8, 9, 10</b>"
    elif fs_div == "APPAREL":
        size_info = "Apparel sizes checked: <b>S, M, L</b>"
    else:
        size_info = "Footwear Men: <b>7,8,9,10</b> &nbsp;·&nbsp; Footwear Women: <b>4,5,6,7</b> &nbsp;·&nbsp; Apparel: <b>S, M, L</b>"

    st.markdown(f"""<div style="background:#f8faff;border:1.5px solid #c7d7f9;border-radius:8px;
        padding:.6rem 1rem;font-size:.8rem;color:#1e40af;margin-bottom:.8rem">
        <b>Cut Size SKU</b> = Article where any expected size has Closing Qty = 0 &nbsp;·&nbsp;
        <b>Full Size SKU</b> = Article where all expected sizes have Closing Qty ≥ 1 &nbsp;·&nbsp;
        {size_info}
    </div>""", unsafe_allow_html=True)

    def get_sz_label(row):
        div = row.get("Division","")
        sz  = str(row.get("Size",""))
        if div == "FOOTWEAR":
            try:
                v = float(sz.replace("½",".5"))
                return str(int(v)) if v == int(v) else str(v)
            except: return None
        elif div == "APPAREL":
            return sz.upper().strip() if sz.upper().strip() in AP_ORDER else None
        return sz

    stk_fs = stk_fs.copy()
    stk_fs["SzLabel"] = stk_fs.apply(get_sz_label, axis=1)
    stk_fs = stk_fs[stk_fs["SzLabel"].notna()]

    # Per Store + Brand + Article → classify
    # Expected sizes
    FW_MEN      = ['7', '8', '9', '10']
    FW_WOMEN    = ['4', '5', '6', '7']
    AP_EXPECTED = ['S', 'M', 'L']

    sku_rows = []
    for (article, brand), grp_df in stk_fs.groupby(["Item ID","Brand"]):
        sizes     = grp_df["SzLabel"].tolist()
        qtys      = grp_df["Closing Qty"].tolist()
        total_qty = sum(qtys)
        div       = grp_df["Division"].iloc[0] if "Division" in grp_df.columns else ""
        gender    = str(grp_df["Gender"].iloc[0]).upper().strip() if "Gender" in grp_df.columns else ""

        # Build size->qty map
        sz_qty = {}
        for s, q in zip(sizes, qtys):
            sz_qty[str(s).strip().upper()] = sz_qty.get(str(s).strip().upper(), 0) + q

        if div == "FOOTWEAR":
            expected = FW_WOMEN if gender == "WOMEN" else FW_MEN
            is_cut = any(sz_qty.get(s, 0) == 0 for s in expected)
        elif div == "APPAREL":
            # Check S, M, L — all three must be present with qty >= 1
            is_cut = any(sz_qty.get(s, 0) == 0 for s in AP_EXPECTED)
        else:
            is_cut = any(q == 0 for q in qtys)

        cls = "✂️ Cut Size" if is_cut else "✅ Full Size"
        stores_list = ", ".join(sorted(grp_df["Store Name"].dropna().unique().tolist()))
        sku_rows.append({
            "Brand": brand, "Article No": article,
            "Gender": gender,
            "Division": div,
            "Stores": stores_list,
            "Store Count": grp_df["Store Name"].nunique(),
            "Total Qty": int(total_qty),
            "Sizes": ", ".join([f"{s}:{int(q)}" for s,q in zip(sizes,qtys)]),
            "Size Count": len(sizes),
            "Classification": cls,
        })

    if not sku_rows:
        st.warning("No data for selected filters.")
    else:
        sku_df  = pd.DataFrame(sku_rows)
        full_df = sku_df[sku_df["Classification"]=="✅ Full Size"].copy()
        cut_df2 = sku_df[sku_df["Classification"]=="✂️ Cut Size"].copy()
        total_sku  = len(sku_df)
        full_count = len(full_df)
        cut_count  = len(cut_df2)
        full_pct   = round(full_count/total_sku*100,1) if total_sku>0 else 0

        # KPI Cards
        k1f,k2f,k3f,k4f = st.columns(4)
        for col,lbl,val,sub,icon,col_bg in [
            (k1f,"Total SKUs",     str(total_sku),  f"Filters: {fs_store} | {fs_brand} | {fs_gender} | {fs_div}",         "📦","linear-gradient(135deg,#6a1b9a,#9c27b0)"),
            (k2f,"Full Size SKUs", str(full_count), f"{full_pct}% of total",           "✅","linear-gradient(135deg,#166534,#16a34a)"),
            (k3f,"Cut Size SKUs",  str(cut_count),  f"{100-full_pct:.1f}% of total",   "✂️","linear-gradient(135deg,#991b1b,#dc2626)"),
            (k4f,"Avg Sizes/SKU",  f"{sku_df['Size Count'].mean():.1f}", "Avg sizes per article","📐","linear-gradient(135deg,#1e40af,#3b82f6)"),
        ]:
            with col:
                st.markdown(f"""<div style="background:{col_bg};border-radius:14px;
                    padding:1rem 1.2rem;box-shadow:0 4px 16px rgba(0,0,0,.15)">
                    <div style="font-size:.58rem;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:rgba(255,255,255,.75)">{icon} {lbl}</div>
                    <div style="font-size:1.6rem;font-weight:800;color:#fff;line-height:1.1">{val}</div>
                    <div style="font-size:.72rem;color:rgba(255,255,255,.7)">{sub}</div>
                </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Charts
        ch1,ch2 = st.columns(2)
        with ch1:
            sec("📊 Full Size vs Cut Size Count")
            fig_fc = go.Figure(go.Bar(
                x=["✅ Full Size","✂️ Cut Size"], y=[full_count,cut_count],
                marker=dict(color=["#16a34a","#dc2626"],line=dict(width=0)),
                text=[f"{full_count} SKUs",f"{cut_count} SKUs"],
                textposition="outside",textfont=dict(size=13,color="#1a0030")))
            fig_fc.update_layout(**cl(320,"Full Size vs Cut Size",margin=dict(l=10,r=10,t=55,b=40)),
                bargap=0.45,yaxis_range=[0,max(full_count,cut_count)*1.28])
            st.plotly_chart(fig_fc, use_container_width=True)

        with ch2:
            sec("🏪 Store-wise Full vs Cut Size")
            sg = sku_df.groupby(["Store","Classification"]).size().unstack(fill_value=0)
            sl = sg.index.tolist()
            fv = [int(sg.loc[s,"✅ Full Size"]) if "✅ Full Size" in sg.columns else 0 for s in sl]
            cv = [int(sg.loc[s,"✂️ Cut Size"])  if "✂️ Cut Size"  in sg.columns else 0 for s in sl]
            fig_sc = go.Figure()
            fig_sc.add_trace(go.Bar(name="✅ Full",x=sl,y=fv,marker_color="#16a34a",text=fv,textposition="outside"))
            fig_sc.add_trace(go.Bar(name="✂️ Cut", x=sl,y=cv,marker_color="#dc2626",text=cv,textposition="outside"))
            fig_sc.update_layout(**cl(320,"Store-wise Full vs Cut",margin=dict(l=10,r=10,t=55,b=110)),
                barmode="group",bargap=0.2,xaxis_tickangle=-45)
            st.plotly_chart(fig_sc, use_container_width=True)

        # Full Size Table
        st.markdown("<br>", unsafe_allow_html=True)
        sec("✅ Full Size SKUs — All expected sizes present (Qty ≥ 1)")
        if len(full_df) > 0:
            fs_show = full_df[["Brand","Gender","Division","Article No","Store Count","Total Qty","Size Count","Sizes"]].sort_values(["Brand","Gender"])
            st.dataframe(
                fs_show.style.apply(lambda x: ["background-color:#dcfce7;color:#166534"]*len(x),axis=1),
                use_container_width=True,hide_index=True)
        else:
            st.info("No Full Size SKUs.")

        # Cut Size Table
        st.markdown("<br>", unsafe_allow_html=True)
        sec("✂️ Cut Size SKUs — Any expected size missing (Qty = 0)")
        if len(cut_df2) > 0:
            cs_show = cut_df2[["Brand","Gender","Division","Article No","Store Count","Total Qty","Size Count","Sizes"]].sort_values(["Brand","Gender"])
            st.dataframe(
                cs_show.style.apply(lambda x: ["background-color:#fee2e2;color:#991b1b"]*len(x),axis=1),
                use_container_width=True,hide_index=True)
        else:
            st.success("✅ No Cut Size SKUs!")

    st.markdown("---")
    sz_div = st.selectbox("Select Division for Size Analysis", DIVISIONS, key="sz_div")
    sale_sz  = sale[sale['Division']==sz_div].copy()
    stock_sz = stock[stock['Division']==sz_div].copy()

    if sz_div == 'FOOTWEAR':
        sale_sz  = sale_sz[sale_sz['Size'].apply(is_fw)].copy()
        stock_sz = stock_sz[stock_sz['Size'].apply(is_fw)].copy()
        sale_sz['SZ']  = sale_sz['Size'].apply(lambda x: float(str(x).replace('½','.5')))
        stock_sz['SZ'] = stock_sz['Size'].apply(lambda x: float(str(x).replace('½','.5')))
        sz_order = sorted(sale_sz['SZ'].unique())
        sz_labels = [str(int(s)) if s==int(s) else str(s) for s in sz_order]
    elif sz_div == 'APPAREL':
        sale_sz  = sale_sz[sale_sz['Size'].apply(is_ap)].copy()
        stock_sz = stock_sz[stock_sz['Size'].apply(is_ap)].copy()
        sale_sz['SZ']  = sale_sz['Size'].str.upper().str.strip().map(AP_MAP).fillna(99)
        stock_sz['SZ'] = stock_sz['Size'].str.upper().str.strip().map(AP_MAP).fillna(99)
        sz_order = sorted([x for x in sale_sz['SZ'].unique() if x!=99])
        sz_labels = [AP_REV.get(int(s),str(s)) for s in sz_order]
    else:
        sale_sz['SZ']  = sale_sz['Size']
        stock_sz['SZ'] = stock_sz['Size']
        sz_order = sorted(sale_sz['SZ'].dropna().unique())
        sz_labels = [str(s) for s in sz_order]

    sz_sq = sale_sz.groupby('SZ')['Sale Qty'].sum().reindex(sz_order).fillna(0)
    sz_kq = stock_sz.groupby('SZ')['Closing Qty'].sum().reindex(sz_order).fillna(0)

    sa6,sb6 = st.columns(2)
    with sa6:
        sec(f"📊 {sz_div} — Sale Qty by Size")
        fig_szs = go.Figure(go.Bar(
            x=sz_labels, y=sz_sq.values,
            marker=dict(color=sz_sq.values,colorscale=BLUE_SEQ,line=dict(width=0)),
            text=[str(int(v)) if v>0 else "" for v in sz_sq.values],
            textposition='outside',textfont=dict(size=11,color='#1a0030')))
        fig_szs.update_layout(**cl(340,f"{sz_div} Sale Qty by Size",margin=dict(l=10,r=10,t=55,b=40)),
            bargap=0.3, yaxis_range=[0,max(sz_sq.max()*1.2,1)])
        st.plotly_chart(fig_szs, use_container_width=True)

    with sb6:
        sec(f"📦 {sz_div} — Stock Qty by Size")
        fig_szk = go.Figure(go.Bar(
            x=sz_labels, y=sz_kq.values,
            marker=dict(color=sz_kq.values,colorscale=[[0,'#fce7f3'],[0.5,'#ec4899'],[1,'#9d174d']],line=dict(width=0)),
            text=[str(int(v)) if v>0 else "" for v in sz_kq.values],
            textposition='outside',textfont=dict(size=11,color='#1a0030')))
        fig_szk.update_layout(**cl(340,f"{sz_div} Stock Qty by Size",margin=dict(l=10,r=10,t=55,b=40)),
            bargap=0.3, yaxis_range=[0,max(sz_kq.max()*1.2,1)])
        st.plotly_chart(fig_szk, use_container_width=True)

    sec(f"📐 {sz_div} — Size Sell-Through Rate")
    st_sz = [round(float(sz_sq.get(s,0))/(float(sz_sq.get(s,0))+float(sz_kq.get(s,0)))*100,1)
             if (float(sz_sq.get(s,0))+float(sz_kq.get(s,0)))>0 else 0 for s in sz_order]
    fig_szst = go.Figure(go.Bar(
        x=sz_labels, y=st_sz,
        marker=dict(color=['#16a34a' if r>=60 else ('#ca8a04' if r>=30 else '#dc2626') for r in st_sz],line=dict(width=0)),
        text=[f"{r:.0f}%" for r in st_sz], textposition='outside', textfont=dict(size=12,color='#1a0030')))
    fig_szst.update_layout(**cl(320,f"{sz_div} Size Sell-Through Rate (%)",margin=dict(l=10,r=10,t=55,b=40)),
        bargap=0.3, yaxis_range=[0,120],
        shapes=[
            dict(type='line',x0=-0.5,x1=len(sz_labels)-0.5,y0=60,y1=60,line=dict(color='#16a34a',width=2,dash='dash')),
            dict(type='line',x0=-0.5,x1=len(sz_labels)-0.5,y0=30,y1=30,line=dict(color='#dc2626',width=2,dash='dash'))
        ])
    st.plotly_chart(fig_szst, use_container_width=True)

    sec(f"✂️ Cut Size Alert — {sz_div}")
    st.markdown("""<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;
        padding:.5rem 1rem;font-size:.8rem;color:#991b1b;margin-bottom:.8rem">
        <b>Cut Size</b> = ✂️ Closing Stock 1–2 pcs only &nbsp;·&nbsp; ❌ Zero Sale in entire period
    </div>""", unsafe_allow_html=True)

    cut_recs = []
    for store in stock_sz['Store Name'].unique():
        s_stk = stock_sz[stock_sz['Store Name']==store]
        s_sal = sale_sz[sale_sz['Store Name']==store] if store in sale_sz['Store Name'].values else pd.DataFrame()
        for sz in s_stk['SZ'].unique():
            qty_k = int(s_stk[s_stk['SZ']==sz]['Closing Qty'].sum())
            val_k = float(s_stk[s_stk['SZ']==sz]['StockValue'].sum())
            qty_s = int(s_sal[s_sal['SZ']==sz]['Sale Qty'].sum()) if len(s_sal)>0 else 0
            cut_type = []
            if 0 < qty_k <= 2: cut_type.append("✂️ Low Qty (1-2 pcs)")
            if qty_s == 0 and qty_k > 0: cut_type.append("❌ Zero Sale")
            if cut_type:
                if sz_div=='FOOTWEAR': slbl = str(int(sz)) if sz==int(sz) else str(sz)
                elif sz_div=='APPAREL': slbl = AP_REV.get(int(sz),str(sz))
                else: slbl = str(sz)
                cut_recs.append({'Store':store,'Size':slbl,'Closing Qty':qty_k,
                    'Stock Value (L)':round(val_k,2),'Sale Qty':qty_s,
                    'Cut Type':' + '.join(cut_type),
                    'Action':'⚠️ Liquidate' if qty_s==0 else '📦 Replenish'})

    if cut_recs:
        cut_df = pd.DataFrame(cut_recs)
        cf = st.selectbox("Filter",["All","✂️ Low Qty (1-2 pcs)","❌ Zero Sale"],key="cut_f")
        if cf != "All": cut_df = cut_df[cut_df['Cut Type'].str.contains(cf.split('(')[0].strip(),na=False)]
        st.dataframe(cut_df.sort_values(['Cut Type','Store']), use_container_width=True, hide_index=True)
        st.markdown(f"""<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;
            padding:.5rem 1rem;font-size:.8rem;color:#991b1b">
            ⚠️ <b>{len(cut_recs)} cut size combinations</b> found in {sz_div}
        </div>""", unsafe_allow_html=True)
    else:
        st.success(f"✅ No cut sizes in {sz_div}!")

    sec("🔍 Filter by Store & Brand")
    fs1,fs2 = st.columns(2)
    with fs1: st_f = st.selectbox("Store",["All"]+sorted(sale_sz['Store Name'].unique()),key="sz_st")
    with fs2: br_f = st.selectbox("Brand",["All"]+sorted(sale_sz['Brand'].unique()),key="sz_br")
    ssf = sale_sz.copy(); skf = stock_sz.copy()
    if st_f!="All": ssf=ssf[ssf['Store Name']==st_f]; skf=skf[skf['Store Name']==st_f]
    if br_f!="All": ssf=ssf[ssf['Brand']==br_f]; skf=skf[skf['Brand']==br_f]
    fsq = ssf.groupby('SZ')['Sale Qty'].sum().reindex(sz_order).fillna(0)
    fkq = skf.groupby('SZ')['Closing Qty'].sum().reindex(sz_order).fillna(0)
    fig_szf = go.Figure()
    fig_szf.add_trace(go.Bar(name='Sale Qty', x=sz_labels,y=fsq.values,marker_color='#7b1fa2'))
    fig_szf.add_trace(go.Bar(name='Stock Qty',x=sz_labels,y=fkq.values,marker_color='#ce93d8'))
    fig_szf.update_layout(**cl(300,f"Size: Sale vs Stock | {st_f} | {br_f}",margin=dict(l=10,r=10,t=55,b=40)),
        barmode='group',bargap=0.2)
    st.plotly_chart(fig_szf, use_container_width=True)

# ══ TAB 7: HEATMAP ══
with t7:
    hm_type = st.selectbox("Heatmap Type",["Store × Category","Store × Brand","Store × Gender","Brand × Category"],key="hm_t")
    if   hm_type=="Store × Category": hm = sale.pivot_table(index='Store Name',columns='Category',values='NetSale',aggfunc='sum').fillna(0)
    elif hm_type=="Store × Brand":    hm = sale.pivot_table(index='Store Name',columns='Brand',   values='NetSale',aggfunc='sum').fillna(0)
    elif hm_type=="Store × Gender":   hm = sale.pivot_table(index='Store Name',columns='Gender',  values='NetSale',aggfunc='sum').fillna(0)
    else:                              hm = sale.pivot_table(index='Brand',     columns='Category',values='NetSale',aggfunc='sum').fillna(0)

    hm_nan = hm.replace(0,np.nan)
    fig_hm = go.Figure(go.Heatmap(
        z=hm_nan.values.tolist(), x=hm_nan.columns.tolist(), y=hm_nan.index.tolist(),
        colorscale=[[0,'#fdf8ff'],[0.2,'#e9d8f8'],[0.5,'#c084fc'],[0.75,'#9333ea'],[1,'#581c87']],
        text=[[fmt_lac(v) if not np.isnan(v) else "—" for v in row] for row in hm_nan.values.tolist()],
        texttemplate="%{text}", textfont=dict(size=9,color='#1a0030'),
        hoverongaps=False,
        colorbar=dict(title="Sale L",tickfont=dict(color='#1a0030'))))
    fig_hm.update_layout(
        paper_bgcolor="rgba(255,255,255,1)", plot_bgcolor="rgba(245,240,255,0.5)",
        font=dict(color="#1a0030",family="Inter",size=11), height=700,
        margin=dict(l=200,r=20,t=55,b=80),
        title=dict(text=f"<b>Sale Heatmap: {hm_type}</b>",font=dict(color='#1a0030',size=14,family='Plus Jakarta Sans')),
        xaxis=dict(tickangle=-30,tickfont=dict(size=10,color='#1a0030')),
        yaxis=dict(tickfont=dict(size=10,color='#1a0030'),autorange='reversed'))
    st.plotly_chart(fig_hm, use_container_width=True)

    if hm.values.max() > 0:
        mi = np.unravel_index(hm.values.argmax(),hm.shape)
        rt = hm.sum(axis=1).sort_values(); ct = hm.sum(axis=0).sort_values(ascending=False)
        st.markdown(f"""<div style="background:#f8faff;border:1.5px solid #c7d7f9;border-radius:12px;padding:1rem 1.2rem;margin-top:.8rem">
          <div style="font-size:.6rem;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#1e40af;margin-bottom:.7rem">🔥 HEATMAP — KEY INSIGHTS</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:.7rem">
            <div style="background:#f5f3ff;border-radius:8px;padding:.7rem;border-left:4px solid #7c3aed">
              <div style="font-size:.65rem;font-weight:700;color:#4c1d95;margin-bottom:.3rem">🏆 HIGHEST COMBINATION</div>
              <div style="font-size:.85rem;font-weight:800;color:#1a0030"><b>{hm.index[mi[0]]}</b> → {hm.columns[mi[1]]}</div>
              <div style="font-size:.78rem;color:#4c1d95">{fmt_lac(hm.values[mi])}</div>
            </div>
            <div style="background:#f0fdf4;border-radius:8px;padding:.7rem;border-left:4px solid #16a34a">
              <div style="font-size:.65rem;font-weight:700;color:#166534;margin-bottom:.3rem">📦 TOP 3 {hm_type.split('×')[1].strip()}S</div>
              <div style="font-size:.78rem;color:#1a0030">{"<br>".join([f"<b>{c}</b> — {fmt_lac(v)}" for c,v in ct.head(3).items()])}</div>
            </div>
            <div style="background:#eff6ff;border-radius:8px;padding:.7rem;border-left:4px solid #1d4ed8">
              <div style="font-size:.65rem;font-weight:700;color:#1e40af;margin-bottom:.3rem">🏆 BEST</div>
              <div style="font-size:.78rem;color:#1a0030"><b>{rt.index[-1]}</b> — {fmt_lac(rt.values[-1])}</div>
            </div>
            <div style="background:#fef2f2;border-radius:8px;padding:.7rem;border-left:4px solid #dc2626">
              <div style="font-size:.65rem;font-weight:700;color:#991b1b;margin-bottom:.3rem">⚠️ WEAKEST</div>
              <div style="font-size:.78rem;color:#1a0030"><b>{rt.index[0]}</b> — {fmt_lac(rt.values[0])}</div>
            </div>
          </div>
        </div>""", unsafe_allow_html=True)

# ══ TAB 8: STORE DEEP DIVE ══
with t8:
    sec("🔍 Store Deep Dive")
    dd_st = st.selectbox("Select Store", sorted(sale['Store Name'].unique()), key="dd_s")
    if dd_st:
        ss = get_store_sale(sale, dd_st)
        sk = get_store_stock(stock, dd_st)
        ts = ss['NetSale'].sum(); tk = sk['StockValue'].sum()
        tq = sk['Closing Qty'].sum()
        rank = int(sale.groupby('Store Name')['NetSale'].sum().rank(ascending=False)[dd_st])
        cont = ts/grand_sale if grand_sale>0 else 0

        m1,m2,m3,m4 = st.columns(4)
        kpi_card(m1,"Total Sale",    fmt_lac(ts),  "Apr'25–Feb'26","💰")
        kpi_card(m2,"Closing Stock", fmt_lac(tk),  f"{int(tq):,} Pcs","📦")
        kpi_card(m3,"Contribution",  pct(cont,4),  "Of total sale","📊")
        kpi_card(m4,"Store Rank",    f"#{rank}",   f"Out of {sale['Store Name'].nunique()} stores","🏅")
        st.markdown("<br>", unsafe_allow_html=True)

        # Month selector
        sel_mon_idx = None
        mon_opts = ["All Months"] + MONTH_SHORT
        sel_mon_label = st.radio("📅 Select Month", mon_opts, horizontal=True, key=f"dd_mon_{dd_st}")
        if sel_mon_label != "All Months":
            sel_mon_idx = MONTH_SHORT.index(sel_mon_label)
            ss_filter = ss[ss['Month'] == MONTHS_ORDER[sel_mon_idx]]
            title_suffix = f" — {sel_mon_label}"
        else:
            ss_filter = ss
            title_suffix = " — All Months"

        d1,d2 = st.columns([3,2])
        with d1:
            mm = ss.groupby('Month')['NetSale'].sum().reindex(MONTHS_ORDER).fillna(0)
            bar_clrs = ['#9c27b0'] * len(MONTH_SHORT)
            if sel_mon_idx is not None:
                bar_clrs = ['#f3e5f5'] * len(MONTH_SHORT)
                bar_clrs[sel_mon_idx] = '#6a1b9a'
            fig_dm = go.Figure(go.Bar(x=MONTH_SHORT, y=mm.values,
                marker=dict(color=bar_clrs, line=dict(width=0)),
                text=[fmt_lac(v) if v>0 else "" for v in mm.values],
                textposition='outside', textfont=dict(size=10, color='#1a0030')))
            fig_dm.update_layout(**cl(280, f"{dd_st} — Monthly Sale", margin=dict(l=10,r=10,t=50,b=40)), bargap=0.3)
            st.plotly_chart(fig_dm, use_container_width=True)

        with d2:
            cd = ss_filter.groupby('Category')['NetSale'].sum()
            cd = cd[cd>0]
            if len(cd)>0:
                fig_dp = go.Figure(go.Pie(labels=cd.index.tolist(), values=cd.values.tolist(), hole=0.48,
                    marker=dict(colors=CAT_COLORS[:len(cd)], line=dict(color='#fff', width=2)),
                    textinfo='label+percent', textfont=dict(size=11, color='#1a0030'),
                    insidetextfont=dict(size=10, color='#fff')))
                fig_dp.update_layout(**cl(280, f"Category Mix{title_suffix}", margin=dict(l=10,r=10,t=50,b=10)))
                st.plotly_chart(fig_dp, use_container_width=True)

        d3,d4 = st.columns(2)
        with d3:
            bd = ss_filter.groupby('Brand')['NetSale'].sum().sort_values(ascending=False)
            bd = bd[bd>0]
            fig_db = go.Figure(go.Bar(x=bd.index.tolist(), y=bd.values.tolist(),
                marker=dict(color=CAT_COLORS[:len(bd)], line=dict(width=0)),
                text=[fmt_lac(v) for v in bd.values], textposition='outside'))
            fig_db.update_layout(**cl(280, f"Brand Mix{title_suffix}", margin=dict(l=10,r=10,t=50,b=60)),
                bargap=0.3, xaxis_tickangle=-30)
            st.plotly_chart(fig_db, use_container_width=True)

        with d4:
            gd = ss_filter.groupby('Gender')['NetSale'].sum().sort_values(ascending=False)
            gd = gd[gd>0]
            fig_dg = go.Figure(go.Pie(labels=gd.index.tolist(), values=gd.values.tolist(), hole=0.48,
                marker=dict(colors=['#7b1fa2','#e91e63','#1565c0','#ff6f00'], line=dict(color='#fff', width=2)),
                textinfo='label+percent', textfont=dict(size=12, color='#1a0030'),
                insidetextfont=dict(size=10, color='#fff')))
            fig_dg.update_layout(**cl(280, f"Gender Mix{title_suffix}", margin=dict(l=10,r=10,t=50,b=10)))
            st.plotly_chart(fig_dg, use_container_width=True)

        if sel_mon_idx is not None:
            sec(f"📋 Full Detail — {sel_mon_label}")
            detail = ss_filter.groupby(['Brand','Category','Gender'])['NetSale'].sum().reset_index()
            detail = detail[detail['NetSale'] > 0].sort_values('NetSale', ascending=False)
            detail['NetSale'] = detail['NetSale'].apply(lambda x: round(x,2))
            detail.columns = ['Brand','Category','Gender','Sale (L)']
            st.dataframe(detail, use_container_width=True, hide_index=True)

        sec("📦 Stock Details")
        sd = sk.groupby(['Division','Category','Brand'])[['Closing Qty','StockValue']].sum()
        sd = sd[sd['Closing Qty']>0].sort_values('StockValue',ascending=False)
        sd['StockValue']  = sd['StockValue'].apply(lambda x: round(x,2) if x!=0 else "")
        sd['Closing Qty'] = sd['Closing Qty'].apply(lambda x: int(x) if x!=0 else "")
        st.dataframe(sd, use_container_width=True)

# ══ TAB 9: PERFORMANCE ══
with t9:
    sec("📊 Month-on-Month Growth")
    mall = sale.groupby('Month')['NetSale'].sum().reindex(MONTHS_ORDER).fillna(0)
    mom  = mall.pct_change()*100
    mp   = [float(v) for v in mom.values[1:]]
    fig_mom = go.Figure(go.Bar(x=MONTH_SHORT[1:],y=mp,
        marker=dict(color=['#16a34a' if v>=0 else '#dc2626' for v in mp],line=dict(width=0)),
        text=[f"{v:+.1f}%" for v in mp],textposition='outside',textfont=dict(size=12,color='#1a0030')))
    fig_mom.update_layout(**cl(340,"MoM Sale Growth (%) — All Stores",margin=dict(l=20,r=20,t=55,b=40)),
        bargap=0.3)
    fig_mom.update_layout(yaxis=dict(gridcolor='#ede9fe',zeroline=True,zerolinecolor='#9c27b0',zerolinewidth=2))
    st.markdown(f"""<div style="background:linear-gradient(135deg,#f5f3ff,#ede9fe);border:1.5px solid #9c27b0;
        border-radius:10px;padding:.5rem 1.1rem;margin-bottom:.5rem;display:inline-block">
        <span style="font-size:.65rem;font-weight:700;color:#6b21a8;text-transform:uppercase;letter-spacing:2px">📌 APR'25 BASE</span>
        <span style="font-size:1rem;font-weight:800;color:#4c1d95;margin-left:1rem">{fmt_lac(mall.values[0])}</span>
    </div>""", unsafe_allow_html=True)
    st.plotly_chart(fig_mom, use_container_width=True)

    p1,p2 = st.columns(2)
    st_tot = sale.groupby('Store Name')['NetSale'].sum()
    with p1:
        sec("🏆 Top 5 Stores")
        t5 = st_tot.nlargest(5).sort_values()
        fig_t5 = go.Figure(go.Bar(x=t5.values,y=t5.index.tolist(),orientation='h',
            marker=dict(color='#16a34a',line=dict(width=0)),
            text=[fmt_lac(v) for v in t5.values],textposition='outside'))
        fig_t5.update_layout(**cl(300,"Top 5 Stores",margin=dict(l=10,r=160,t=40,b=20)),
            xaxis_range=[0,t5.max()*1.45])
        st.plotly_chart(fig_t5, use_container_width=True)

    with p2:
        sec("🔴 Bottom 5 Stores")
        b5 = st_tot.nsmallest(5).sort_values(ascending=False)
        fig_b5 = go.Figure(go.Bar(x=b5.values,y=b5.index.tolist(),orientation='h',
            marker=dict(color='#dc2626',line=dict(width=0)),
            text=[fmt_lac(v) for v in b5.values],textposition='outside'))
        fig_b5.update_layout(**cl(300,"Bottom 5 Stores",margin=dict(l=10,r=160,t=40,b=20)),
            xaxis_range=[0,b5.max()*1.55])
        st.plotly_chart(fig_b5, use_container_width=True)

    sec("❌ Zero Sale Months — Store-wise")
    swc9 = sale.pivot_table(index='Store Name',columns='Month',values='NetSale',aggfunc='sum').reindex(columns=MONTHS_ORDER).fillna(0)
    zr = []
    for sn in swc9.index:
        row = swc9.loc[sn]
        zm = [MONTH_SHORT[i] for i,v in enumerate(row.values) if v==0]
        sm = [MONTH_SHORT[i] for i,v in enumerate(row.values) if v>0]
        if zm:
            last=float(row.values[-1]); prev=float(row.values[-2])
            g = f"{'▲' if last>=prev else '▼'} {abs((last-prev)/prev*100):.1f}%" if prev>0 else "N/A"
            zr.append({'Store':sn,'Sale Months':', '.join(sm),'Zero Months':', '.join(zm),
                       'Zero Count':len(zm),'Jan→Feb':g,'Total Sale':fmt_lac(st_tot.get(sn,0))})
    if zr:
        st.dataframe(pd.DataFrame(zr).sort_values('Zero Count',ascending=False),
            use_container_width=True, hide_index=True)

    sec("🏷️ Brand Jan→Feb Growth")
    bwc9 = sale.pivot_table(index='Brand',columns='Month',values='NetSale',aggfunc='sum').reindex(columns=MONTHS_ORDER).fillna(0)
    bg = ((bwc9.iloc[:,-1]-bwc9.iloc[:,-2])/bwc9.iloc[:,-2]*100).replace([np.inf,-np.inf],np.nan)
    fig_bg = go.Figure(go.Bar(x=bg.index.tolist(),y=bg.values.tolist(),
        marker=dict(color=['#16a34a' if (v>=0 and not np.isnan(v)) else '#dc2626' for v in bg.values],line=dict(width=0)),
        text=[f"{v:+.1f}%" if pd.notna(v) else "N/A" for v in bg.values],
        textposition='outside',textfont=dict(size=11,color='#1a0030')))
    fig_bg.update_layout(**cl(300,"Brand Jan→Feb Growth (%)",margin=dict(l=10,r=10,t=55,b=60)),
        bargap=0.3,xaxis_tickangle=-30)
    st.plotly_chart(fig_bg, use_container_width=True)

# ══ TAB 10: INVENTORY ══
with t10:
    sec("📊 Sell-Through Rate — Store × Category")
    st.markdown("""<div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;
        padding:.5rem 1rem;font-size:.8rem;color:#166534;margin-bottom:.8rem">
        <b>Sell-Through</b> = Sale ÷ (Sale + Closing Stock) × 100 &nbsp;·&nbsp;
        🟢 ≥60% Good &nbsp;·&nbsp; 🟡 30–59% Avg &nbsp;·&nbsp; 🔴 &lt;30% Slow
    </div>""", unsafe_allow_html=True)

    cats10 = sorted(sale['Category'].dropna().unique())
    strs10 = sorted(sale['Store Name'].dropna().unique())
    stm = pd.DataFrame(index=strs10, columns=cats10, dtype=float)
    for s10 in strs10:
        for c10 in cats10:
            sv = float(sale[(sale['Store Name']==s10)&(sale['Category']==c10)]['NetSale'].sum())
            kv = float(stock[(stock['Store Name']==s10)&(stock['Category']==c10)]['StockValue'].sum())
            t  = sv+kv
            stm.loc[s10,c10] = round(sv/t*100,1) if t>0 else 0
    stm = stm.fillna(0).astype(float)

    fig_stm = go.Figure(go.Heatmap(
        z=stm.values.tolist(), x=stm.columns.tolist(), y=stm.index.tolist(),
        colorscale=[[0,'#fef2f2'],[0.3,'#fca5a5'],[0.6,'#fde68a'],[0.8,'#86efac'],[1,'#16a34a']],
        text=[[f"{v:.0f}%" if v>0 else "—" for v in row] for row in stm.values.tolist()],
        texttemplate="%{text}", textfont=dict(size=9,color='#1a0030'),
        hoverongaps=False, zmin=0, zmax=100,
        colorbar=dict(title="ST%",tickfont=dict(color='#1a0030',size=9),ticksuffix="%")))
    fig_stm.update_layout(
        paper_bgcolor="rgba(255,255,255,1)", plot_bgcolor="rgba(255,255,255,1)",
        font=dict(color="#1a0030",family="Inter",size=11), height=700,
        margin=dict(l=200,r=30,t=50,b=80),
        title=dict(text="<b>Sell-Through Rate (%) — Store × Category</b>",
                   font=dict(color='#1a0030',size=14,family='Plus Jakarta Sans')),
        xaxis=dict(tickangle=-30,tickfont=dict(size=10,color='#1a0030')),
        yaxis=dict(tickfont=dict(size=10,color='#1a0030'),autorange='reversed'))
    st.plotly_chart(fig_stm, use_container_width=True)

    avg_st = float(stm.replace(0,np.nan).stack().mean())
    sa10 = stm.replace(0,np.nan).mean(axis=1).sort_values(ascending=False)
    best_st  = sa10[sa10>=60].index.tolist()[:3]
    worst_st = sa10[sa10<30].index.tolist()[:3]

    restock = []; dead = []
    for s10 in strs10:
        for c10 in cats10:
            sv = float(sale[(sale['Store Name']==s10)&(sale['Category']==c10)]['NetSale'].sum())
            kv = float(stock[(stock['Store Name']==s10)&(stock['Category']==c10)]['StockValue'].sum())
            t  = sv+kv
            if t>0 and sv>5 and sv/t*100>=80: restock.append(f"{s10} → {c10}")
            if sv==0 and kv>0: dead.append(f"{s10} → {c10}")

    dead_val = sum([float(stock[(stock['Store Name']==r.split(' → ')[0])&(stock['Category']==r.split(' → ')[1])]['StockValue'].sum()) for r in dead])

    st.markdown(f"""<div style="background:#f8faff;border:1.5px solid #c7d7f9;border-radius:12px;
        padding:1rem 1.2rem;margin-bottom:1.2rem">
      <div style="font-size:.6rem;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#1e40af;margin-bottom:.7rem">📊 KEY INSIGHTS</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:.7rem">
        <div style="background:#f0fdf4;border-radius:8px;padding:.7rem;border-left:4px solid #16a34a">
          <div style="font-size:.65rem;font-weight:700;color:#166534;margin-bottom:.3rem">🏆 BEST STORES (ST≥60%)</div>
          <div style="font-size:.78rem;color:#1a0030">{"<br>".join([f"<b>{s}</b>" for s in best_st]) if best_st else "N/A"}</div>
        </div>
        <div style="background:#fef2f2;border-radius:8px;padding:.7rem;border-left:4px solid #dc2626">
          <div style="font-size:.65rem;font-weight:700;color:#991b1b;margin-bottom:.3rem">⚠️ NEEDS ATTENTION (ST&lt;30%)</div>
          <div style="font-size:.78rem;color:#1a0030">{"<br>".join([f"<b>{s}</b>" for s in worst_st]) if worst_st else "All OK"}</div>
        </div>
        <div style="background:#eff6ff;border-radius:8px;padding:.7rem;border-left:4px solid #1d4ed8">
          <div style="font-size:.65rem;font-weight:700;color:#1e40af;margin-bottom:.3rem">📦 RESTOCK URGENT ({len(restock)} items)</div>
          <div style="font-size:.78rem;color:#1a0030">{"<br>".join(restock[:4]) if restock else "None"}</div>
        </div>
        <div style="background:#fefce8;border-radius:8px;padding:.7rem;border-left:4px solid #ca8a04">
          <div style="font-size:.65rem;font-weight:700;color:#854d0e;margin-bottom:.3rem">⚠️ DEAD STOCK ({len(dead)} items · {fmt_lac(dead_val)})</div>
          <div style="font-size:.78rem;color:#1a0030">{"<br>".join(dead[:4]) if dead else "None"}</div>
        </div>
      </div>
      <div style="margin-top:.7rem;padding:.5rem .7rem;background:#fff;border-radius:6px;font-size:.75rem;color:#374151">
        📌 <b>Overall Avg ST: {avg_st:.1f}%</b>
      </div>
    </div>""", unsafe_allow_html=True)

    sec("📋 Stock Recommendations")
    recs = []
    for s10 in strs10:
        for c10 in cats10:
            sv = float(sale[(sale['Store Name']==s10)&(sale['Category']==c10)]['NetSale'].sum())
            kv = float(stock[(stock['Store Name']==s10)&(stock['Category']==c10)]['StockValue'].sum())
            kq = int(stock[(stock['Store Name']==s10)&(stock['Category']==c10)]['Closing Qty'].sum())
            t  = sv+kv
            if t==0: continue
            r = sv/t*100
            if   r>=75 and sv>5:  rec="📦 Increase Stock — High Demand"; pri="🔴 Urgent"
            elif r>=60 and sv>0:  rec="📦 Replenish — Good Seller";       pri="🟡 Medium"
            elif r<20 and kv>10:  rec="🔄 Transfer to Better Store";      pri="🔴 Urgent"
            elif sv==0 and kv>0:  rec="❌ Remove — Dead Stock";            pri="🔴 Urgent"
            elif r<30 and kv>5:   rec="⬇️ Reduce Stock — Slow Mover";     pri="🟡 Medium"
            else: continue
            recs.append({'Store':s10,'Category':c10,'Sale (L)':round(sv,2),'Stock (L)':round(kv,2),
                         'Stock Qty':kq,'ST%':f"{r:.1f}%",'Recommendation':rec,'Priority':pri})

    if recs:
        rd = pd.DataFrame(recs)
        pf = st.selectbox("Filter Priority",["All","🔴 Urgent","🟡 Medium"],key="rec_p")
        if pf!="All": rd = rd[rd['Priority']==pf]
        st.dataframe(rd, use_container_width=True, hide_index=True)

# ══ TAB 11: AI STRATEGY ══
with t11:
    sec("🤖 AI Strategy Summary")
    st.markdown("""<div style="background:linear-gradient(135deg,#3a0068,#6a1b9a);border-radius:12px;
        padding:1rem 1.4rem;margin-bottom:1rem;color:#fff">
        <div style="font-size:.65rem;font-weight:700;letter-spacing:2px;text-transform:uppercase;
            color:rgba(255,255,255,.7);margin-bottom:.3rem">HOW IT WORKS</div>
        <div style="font-size:.85rem">AI analyses stores, brands, categories, genders, sizes,
        sell-through and dead stock — generates smart action plan.</div>
    </div>""", unsafe_allow_html=True)

    mall_ai  = sale.groupby('Month')['NetSale'].sum().reindex(MONTHS_ORDER).fillna(0)
    brand_ai = sale.groupby('Brand')['NetSale'].sum().sort_values(ascending=False)
    cat_ai   = sale.groupby('Category')['NetSale'].sum().sort_values(ascending=False)
    gdr_ai   = sale.groupby('Gender')['NetSale'].sum().sort_values(ascending=False)
    t3s      = sale.groupby('Store Name')['NetSale'].sum().nlargest(3)
    b3s      = sale.groupby('Store Name')['NetSale'].sum().nsmallest(3)
    mom_last = ((mall_ai.values[-1]-mall_ai.values[-2])/mall_ai.values[-2]*100) if mall_ai.values[-2]>0 else 0

    prompt = f"""You are a senior retail consultant for SS Retail (UCB multi-brand stores, India).

SALES DATA (Apr'25–Feb'26, 11 months, values in Lacs INR):
- Total Net Sale: {fmt_lac(grand_sale)} across {len(all_stores)} stores
- Total Closing Stock: {fmt_lac(grand_stock)} ({int(grand_qty):,} pcs)
- Jan→Feb Growth: {mom_last:+.1f}%
- Best Month: {MONTH_SHORT[int(mall_ai.values.argmax())]} | Worst: {MONTH_SHORT[int(mall_ai.values.argmin())]}

BRAND PERFORMANCE:
{chr(10).join([f"- {b}: {fmt_lac(v)} ({v/grand_sale*100:.1f}%)" for b,v in brand_ai.items()])}

CATEGORY PERFORMANCE:
{chr(10).join([f"- {c}: {fmt_lac(v)}" for c,v in cat_ai.items()])}

GENDER SPLIT:
{chr(10).join([f"- {g}: {fmt_lac(v)} ({v/grand_sale*100:.1f}%)" for g,v in gdr_ai.items()])}

TOP 3 STORES: {" | ".join([f"{s}: {fmt_lac(v)}" for s,v in t3s.items()])}
BOTTOM 3 STORES: {" | ".join([f"{s}: {fmt_lac(v)}" for s,v in b3s.items()])}

INVENTORY: Dead Stock={len(dead)} combos | Restock Needed={len(restock)} combos | Avg ST={avg_st:.1f}%

Provide strategy in these EXACT sections:
1. EXECUTIVE SUMMARY
2. KEY STRENGTHS
3. CRITICAL ISSUES
4. HOW TO INCREASE SALE — 5 STRATEGIES
5. INVENTORY ACTION PLAN
6. IMMEDIATE PRIORITIES (This Week / This Month / Next Quarter)

Be specific with store/brand names. Direct and actionable."""

    a1,a2,a3 = st.columns([1,2,1])
    with a2:
        gen_ai = st.button("🤖  Generate AI Strategy", use_container_width=True)

    if gen_ai:
        st.session_state.ai_loading = True
        st.session_state.ai_text    = None

    if st.session_state.ai_loading and st.session_state.ai_text is None:
        with st.spinner("🤖 AI data analyse kar raha hai..."):
            try:
                import os
                groq_key = os.environ.get("GROQ_API_KEY", "")
                if not groq_key:
                    st.session_state.ai_text = "❌ GROQ_API_KEY not set in Streamlit Secrets"
                    st.session_state.ai_loading = False
                else:
                    resp = requests.post(
                        "https://api.groq.com/openai/v1/chat/completions",
                        headers={
                            "Authorization": f"Bearer {groq_key}",
                            "Content-Type": "application/json"
                        },
                        json={
                            "model": "llama-3.3-70b-versatile",
                            "messages": [{"role":"user","content": prompt}],
                            "max_tokens": 2000,
                            "temperature": 0.7
                        },
                        timeout=60
                    )
                    if resp.status_code == 200:
                        st.session_state.ai_text = resp.json()['choices'][0]['message']['content']
                        st.session_state.ai_loading = False
                    else:
                        st.session_state.ai_text = f"❌ API Error {resp.status_code}: {resp.text[:300]}"
                        st.session_state.ai_loading = False
            except Exception as e:
                st.session_state.ai_text = f"❌ Error: {str(e)}"
                st.session_state.ai_loading = False
        st.rerun()

    if st.session_state.ai_text:
        txt = st.session_state.ai_text
        if txt.startswith("❌"):
            st.error(txt)
        else:
            section_styles = {
                "EXECUTIVE SUMMARY":    ("📋","#1e3a5f","#eff6ff","#1e40af"),
                "KEY STRENGTHS":        ("💪","#166534","#f0fdf4","#16a34a"),
                "CRITICAL ISSUES":      ("🚨","#991b1b","#fef2f2","#dc2626"),
                "HOW TO INCREASE SALE": ("🚀","#4c1d95","#f5f3ff","#7c3aed"),
                "INVENTORY ACTION PLAN":("📦","#854d0e","#fefce8","#ca8a04"),
                "IMMEDIATE PRIORITIES": ("⚡","#065f46","#ecfdf5","#059669"),
            }
            lines = txt.split('\n')
            cur = None; secs = {}; cont = []
            for line in lines:
                line = line.strip()
                if not line: continue
                found = False
                for k in section_styles:
                    if k in line.upper():
                        if cur: secs[cur] = '\n'.join(cont)
                        cur=k; cont=[]; found=True; break
                if not found and cur: cont.append(line)
            if cur: secs[cur] = '\n'.join(cont)

            for k,(icon,tc,bg,bc) in section_styles.items():
                if k in secs:
                    fmt = []
                    for l in secs[k].split('\n'):
                        l = l.strip()
                        if not l: continue
                        if l.startswith(('-','•','*')): l = '• ' + l.lstrip('-•* ').strip()
                        fmt.append(f'<div style="margin:.25rem 0;font-size:.85rem;color:#1a0030;line-height:1.5">{l}</div>')
                    st.markdown(f"""<div style="background:{bg};border-left:4px solid {bc};
                        border-radius:10px;padding:.9rem 1.1rem;margin-bottom:.8rem">
                        <div style="font-size:.6rem;font-weight:800;letter-spacing:2px;text-transform:uppercase;
                            color:{tc};margin-bottom:.5rem">{icon} {k}</div>
                        {''.join(fmt)}
                    </div>""", unsafe_allow_html=True)

            x1,x2,x3 = st.columns([1,2,1])
            with x2:
                export = f"SS RETAIL — AI STRATEGY\nGenerated: {pd.Timestamp.now().strftime('%d %b %Y %I:%M %p')}\n\n{txt}"
                st.download_button("📥 Download Strategy (TXT)", data=export,
                    file_name="SS_Strategy.txt", mime="text/plain", use_container_width=True)
    elif not st.session_state.ai_loading:
        st.markdown("""<div style="text-align:center;padding:3rem 0">
          <div style="font-size:3rem">🤖</div>
          <div style="font-size:1rem;color:#607d9b;font-weight:500;margin-top:.8rem">
            Generate button click karo — AI analysis tayaar ho jaayegi</div>
          <div style="font-size:.8rem;color:#90a4c0;margin-top:.4rem">
            Stores · Brands · Categories · Sizes · Inventory sab analyse hoga</div>
        </div>""", unsafe_allow_html=True)
