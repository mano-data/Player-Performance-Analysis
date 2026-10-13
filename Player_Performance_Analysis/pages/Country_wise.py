import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import openpyxl, base64, os, math

# ── PAGE CONFIG ───────────────────────────────────────────────
st.set_page_config(
    page_title="Country Wise | Cricket Analysis",
    page_icon="", layout="wide",
    initial_sidebar_state="collapsed",
)

# ── PATHS ─────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR    = os.path.dirname(BASE_DIR)
DATASET_DIR = os.path.join(ROOT_DIR, "Dataset")
IMAGE_DIR   = os.path.join(ROOT_DIR, "images")

def _find(fname):
    for d in [DATASET_DIR, ROOT_DIR, BASE_DIR]:
        p = os.path.join(d, fname)
        if os.path.exists(p): return p
    return os.path.join(DATASET_DIR, fname)

MASTER_PATH = _find("player_master.csv")
STATS_PATH  = _find("player_stats2.xlsx")

# ── PALETTE ───────────────────────────────────────────────────
NAVY  = "#1B3A6B"
AMBER = "#E8A020"
STEEL = "#4A6FA5"
GREEN = "#1E7A34"
RED   = "#C0392B"
BORDER= "#dde3ee"

# Country color mapping
CTR_COLORS = {
    "India":        "#FF6600",
    "Australia":    "#F5A623",
    "England":      "#003087",
    "South Africa": "#007A4D",
    "New Zealand":  "#7F8C8D",
    "West Indies":  "#6B3A2A",
    "Pakistan":     "#1A5C2A",
    "Sri Lanka":    "#0066CC",
    "Bangladesh":   "#E74C3C",
    "Afghanistan":  "#00205B",
    "UAE":          "#E55812",
    "default":      NAVY,
}
SENA = {"South Africa","England","New Zealand","Australia"}
ASIA = {"India","Sri Lanka","Bangladesh","Pakistan","UAE"}

def cclr(name):
    return CTR_COLORS.get(str(name).strip(), CTR_COLORS["default"])

def cclrs(series):
    return [cclr(n) for n in series]

CHART_COLORS = [NAVY, AMBER, STEEL, GREEN, RED, "#7B5EA7","#20B2AA","#E55812","#3A9BD5","#F39C12"]

# ── CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700;800&display=swap');

html,body,[class*="css"]{font-family:'Inter',sans-serif;background:#fff !important;}
#MainMenu,footer,header{visibility:hidden;}
section[data-testid="stSidebar"]{display:none !important;}

.block-container,
[data-testid="stAppViewBlockContainer"],
[data-testid="block-container"]{
    padding:0 60px 60px 60px !important;
    max-width:1560px !important;
    margin:0 auto !important;
    background:#fff !important;
}
[data-testid="stAppViewBlockContainer"]>div:first-child{padding-top:0 !important;}

/* NAVBAR */
.top-header{
    background:#fff; border-bottom:2px solid #1B3A6B;
    padding:0 60px; height:64px;
    display:flex; align-items:center; justify-content:space-between;
    position:fixed;width:100%; top:0; z-index:999;
    box-shadow:0 2px 12px rgba(27,58,107,.08);
    margin:0 -60px 0 -60px;
}
.header-brand{font-family:'Poppins',sans-serif;font-size:1.2rem;font-weight:700;color:#1B3A6B;}
.header-brand span{color:#E8A020;}
.header-nav{display:flex;gap:8px;align-items:center;}
.h-btn{background:#1B3A6B;color:#fff !important;border:none;padding:7px 18px;border-radius:6px;
    font-size:.82rem;font-weight:600;text-decoration:none !important;display:inline-block;}
.h-btn:hover{background:#243f7a;}
.h-btn-ghost{background:transparent;color:#1B3A6B !important;border:1.5px solid #1B3A6B;
    padding:6px 16px;border-radius:6px;font-size:.82rem;font-weight:600;
    text-decoration:none !important;display:inline-block;}
.h-btn-ghost:hover{background:#f0f4fb;}

/* PAGE TITLE */
.page-title-bar{background:#f7f9fc;border-left:4px solid #1B3A6B;
    padding:14px 60px;margin:40px -60px 28px -60px;}
.page-title-bar h2{font-family:'Poppins',sans-serif;font-size:1.15rem;font-weight:700;color:#1B3A6B;margin:0;}
.page-title-bar p{color:#6B7A8D;font-size:.8rem;margin:20px 0 0 0;}

/* FILTER */
.filter-wrap{background:#f7f9fc;border:1px solid #dde3ee;border-radius:10px;
    padding:22px 28px;margin-bottom:26px;}
.filter-label{font-size:.75rem;font-weight:600;color:#1B3A6B;
    letter-spacing:.8px;text-transform:uppercase;margin-bottom:6px;}
.filter-note{font-size:.72rem;color:#6B7A8D;margin-top:8px;}

.stSelectbox>div>div{background:#fff !important;border:1.5px solid #c8d0e0 !important;
    border-radius:7px !important;color:#1B3A6B !important;font-size:.88rem !important;}
.stSelectbox>div>div:focus-within{border-color:#1B3A6B !important;
    box-shadow:0 0 0 3px rgba(27,58,107,.1) !important;}
div[data-testid="stButton"]>button{background:#1B3A6B !important;color:#fff !important;
    font-family:'Poppins',sans-serif !important;font-weight:700 !important;
    font-size:.9rem !important;letter-spacing:1px !important;border:none !important;
    border-radius:7px !important;padding:12px 0 !important;width:100% !important;}
div[data-testid="stButton"]>button:hover{background:#243f7a !important;}

/* SECTION TITLE */
.sec-title{font-family:'Poppins',sans-serif;font-size:1rem;font-weight:700;color:#1B3A6B;
    margin:0 0 14px 0;padding-bottom:8px;border-bottom:2px solid #e8edf5;
    display:flex;align-items:center;gap:8px;}
.sec-title .dot{width:8px;height:8px;background:#1B3A6B;border-radius:50%;display:inline-block;}

/* PROFILE CARD */
.profile-card{background:#fff;border:1px solid #dde3ee;border-radius:12px;
    padding:28px 22px;box-shadow:0 2px 12px rgba(27,58,107,.07);
    position:relative;overflow:hidden;}
.profile-card::before{content:'';position:absolute;top:0;left:0;right:0;height:4px;
    background:linear-gradient(90deg,#1B3A6B 0%,#E8A020 100%);}
.player-name{font-family:'Poppins',sans-serif;font-size:1.2rem;font-weight:700;
    color:#1B3A6B;text-align:center;margin:12px 0 4px 0;}
.player-role-badge{display:block;width:fit-content;margin:0 auto 18px auto;
    padding:3px 14px;border-radius:20px;font-size:.72rem;font-weight:700;
    letter-spacing:.5px;text-transform:uppercase;}
.badge-batsman{background:#EBF3FF;color:#1B3A6B;border:1px solid #c0d0eb;}
.badge-bowler{background:#FFF0E8;color:#C0501A;border:1px solid #f0c0a0;}
.badge-allrounder{background:#F0F8EE;color:#2A7A2A;border:1px solid #a8d4a0;}
.info-row{display:flex;justify-content:space-between;align-items:center;
    padding:9px 0;border-bottom:1px solid #f0f3f8;font-size:.83rem;}
.info-row:last-child{border-bottom:none;}
.info-lbl{color:#6B7A8D;font-weight:500;}
.info-val{color:#1B3A6B;font-weight:600;}

/* COUNTRY BANNER */
.ctr-banner{background:linear-gradient(135deg,#1B3A6B,#243f7a);border-radius:10px;
    padding:16px 22px;display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;}
.ctr-banner-lbl{color:rgba(255,255,255,.65);font-size:.74rem;text-transform:uppercase;letter-spacing:1px;}
.ctr-banner-val{font-family:'Poppins',sans-serif;font-size:1.3rem;font-weight:700;color:#fff;margin-top:2px;}
.ctr-banner-sub{color:rgba(255,255,255,.6);font-size:.78rem;margin-top:4px;}
.ctr-banner-pct{font-family:'Poppins',sans-serif;font-size:1.6rem;font-weight:700;color:#E8A020;}
.ctr-banner-pct-lbl{color:rgba(255,255,255,.65);font-size:.72rem;text-transform:uppercase;letter-spacing:.8px;}

/* INSIGHT CARDS */
.insight-card{background:#f7f9fc;border:1px solid #dde3ee;border-left:4px solid #E8A020;
    border-radius:8px;padding:12px 16px;display:flex;align-items:flex-start;
    gap:12px;margin-bottom:9px;}
.insight-icon{font-size:1.1rem;flex-shrink:0;margin-top:2px;}
.insight-text{font-size:.84rem;color:#2c3e50;line-height:1.5;}
.insight-text strong{color:#1B3A6B;}

/* KPI CARDS */
.kpi-card{background:#fff;border:1px solid #dde3ee;border-radius:10px;
    padding:18px 14px 14px 14px;text-align:center;
    box-shadow:0 2px 8px rgba(27,58,107,.05);
    transition:box-shadow .2s,transform .2s;position:relative;overflow:hidden;}
.kpi-card:hover{box-shadow:0 6px 18px rgba(27,58,107,.12);transform:translateY(-2px);}
.kpi-card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;background:#1B3A6B;}
.kpi-card.hl::after{background:#E8A020;}
.kpi-card.hl-g::after{background:#1E7A34;}
.kpi-val{font-family:'Poppins',sans-serif;font-size:1.65rem;font-weight:700;
    color:#1B3A6B;line-height:1;margin-bottom:6px;}
.kpi-lbl{font-size:.7rem;color:#8899aa;font-weight:600;letter-spacing:.6px;text-transform:uppercase;}
.kpi-sub{font-size:.68rem;color:#aab0bc;margin-top:3px;}

/* SCENARIO BANNER */
.sc-banner{background:#f7f9fc;border:1px solid #dde3ee;border-radius:8px;
    padding:13px 20px;margin-bottom:22px;display:flex;align-items:center;gap:14px;}
.sc-pill{background:#1B3A6B;color:#fff;padding:4px 14px;border-radius:20px;
    font-size:.72rem;font-weight:700;letter-spacing:.5px;white-space:nowrap;}
.sc-desc{font-size:.83rem;color:#6B7A8D;}

/* DIVIDER / NO DATA */
.divider{border:none;height:1px;background:#e8edf5;margin:28px 0;}
.no-data{text-align:center;padding:50px;color:#8899aa;font-size:.9rem;}
.no-data .icon{font-size:2.5rem;display:block;margin-bottom:10px;}

/* ACTION BAR + FOOTER */
.action-bar{background:#f7f9fc;border-top:1px solid #dde3ee;
    padding:26px 60px;margin:32px -60px 0 -60px;
    display:flex;align-items:center;justify-content:space-between;}
.act-btn{display:inline-flex;align-items:center;gap:8px;padding:11px 28px;border-radius:7px;
    font-family:'Poppins',sans-serif;font-weight:700;font-size:.88rem;letter-spacing:.5px;
    cursor:pointer;text-decoration:none !important;border:none;}
.act-report{background:#fff;color:#1B3A6B !important;border:2px solid #1B3A6B;}
.act-report:hover{background:#EBF3FF;}
.act-compare{background:#1B3A6B;color:#fff !important;}
.act-compare:hover{background:#243f7a;}
.footer{background:#1B3A6B;color:rgba(255,255,255,.7);text-align:center;
    padding:18px 60px;font-size:.8rem;margin:0 -60px -60px -60px;letter-spacing:.3px;}
.footer span{color:#E8A020;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ── DATA LOADING ──────────────────────────────────────────────
@st.cache_data
def load_data():
    master = pd.read_csv(MASTER_PATH)
    master.columns = master.columns.str.strip()
    wb   = openpyxl.load_workbook(STATS_PATH)
    skip = {"Match_id","Opponent","Venue","Country","Match_result",
            "Dismissed_method","Dismissed_Length","Dismissed_against",
            "Dismissed_by","Dismissed _by","Match_date"}
    stats = {}
    for sheet in wb.sheetnames:
        ws   = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            stats[sheet.strip()] = pd.DataFrame(); continue
        cols = [str(c).strip() if c else f"_c{i}" for i, c in enumerate(rows[0])]
        df   = pd.DataFrame(rows[1:], columns=cols)
        df["Match_date"] = pd.to_datetime(df["Match_date"], errors="coerce")
        for c in df.columns:
            if c not in skip:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        stats[sheet.strip()] = df
    return master, stats

try:
    master_df, stats_dict = load_data()
except Exception as e:
    st.error(f"Data load error: {e}"); st.stop()

# ── HELPERS ───────────────────────────────────────────────────
def get_img_b64(filename):
    for d in [IMAGE_DIR, os.path.join(ROOT_DIR,"images"), BASE_DIR]:
        p = os.path.join(d, filename)
        if os.path.exists(p):
            ext  = filename.rsplit(".",1)[-1].lower()
            mime = "jpeg" if ext in ["jpg","jpeg"] else ext
            with open(p,"rb") as f:
                return f"data:image/{mime};base64,{base64.b64encode(f.read()).decode()}"
    return None

def initials_avatar(name):
    ini = "".join(p[0].upper() for p in name.split()[:2])
    return (f'<div style="width:100px;height:100px;border-radius:50%;background:#EBF3FF;'
            f'border:3px solid #1B3A6B;display:flex;align-items:center;justify-content:center;'
            f'margin:0 auto;font-family:Poppins,sans-serif;font-size:1.8rem;font-weight:700;'
            f'color:#1B3A6B;">{ini}</div>')

def calc_avg(runs_s, not_out_s, innings):
    tot = float(runs_s.sum())
    dis = max(int(innings) - int(not_out_s.sum()), 0)
    return round(tot / dis, 2) if dis > 0 else round(tot, 2)

def pchart(fig, h=310, title=""):
    """Consistent chart styling."""
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Inter", color="#2c3e50", size=11),
        height=h, margin=dict(l=10,r=10,t=44,b=10),
        title=dict(text=title, font=dict(family="Poppins",size=13,color=NAVY), x=0.02, xanchor="left"),
        legend=dict(bgcolor="rgba(255,255,255,.9)", bordercolor=BORDER, borderwidth=1, font=dict(size=10)),
    )
    fig.update_xaxes(gridcolor="#f0f3f8", linecolor=BORDER, tickfont=dict(size=9), zeroline=False)
    fig.update_yaxes(gridcolor="#f0f3f8", linecolor=BORDER, tickfont=dict(size=9), zeroline=False)
    return fig

def kcard(val, lbl, sub="", v=""):
    return (f'<div class="kpi-card {v}"><div class="kpi-val">{val}</div>'
            f'<div class="kpi-lbl">{lbl}</div>'
            + (f'<div class="kpi-sub">{sub}</div>' if sub else "") + '</div>')

def ins(icon, txt):
    return (f'<div class="insight-card"><span class="insight-icon">{icon}</span>'
            f'<span class="insight-text">{txt}</span></div>')

def get_col(df, *names):
    for n in names:
        if n in df.columns: return df[n]
    return pd.Series(0, index=df.index)

def bat_rows(df):
    return df[df["Batted_Flag"]==1] if "Batted_Flag" in df.columns else df

def bowl_rows(df):
    return df[df["Bowl_Flag"]==1] if "Bowl_Flag" in df.columns else df

def pch(fig):
    st.plotly_chart(fig, use_container_width=True)

# ── NAVBAR ────────────────────────────────────────────────────
st.markdown("""
<div class="top-header">
  <div class="header-brand">Cricket Player<span> Performance Analysis</span></div>
  <div class="header-nav">
    <a class="h-btn" href="/" target="_self">Back</a>
    <a class="h-btn" href="/userhome" target="_self">Home</a>
    <a class="h-btn" href="/admin" target="_self">Admin</a>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="page-title-bar">
  <h2>Country Wise Performance Analysis</h2>
  <p>Analyse performance across countries and dive into venue-level insights when a specific country is selected</p>
</div>
""", unsafe_allow_html=True)

# ── FILTER PANEL ──────────────────────────────────────────────
fc1, fc2, fc3 = st.columns([3, 3, 2])

with fc1:
    st.markdown('<div class="filter-label"> Select Player</div>', unsafe_allow_html=True)
    sel_player = st.selectbox("Player", sorted(stats_dict.keys()), label_visibility="collapsed")

_raw = stats_dict.get(sel_player, pd.DataFrame())

with fc2:
    st.markdown('<div class="filter-label"> Select Opponent</div>', unsafe_allow_html=True)

    odi_nations = [
        "Afghanistan","Australia","Bangladesh","England","Ireland",
        "New Zealand","Pakistan","South Africa","Sri Lanka","West Indies",
        "Zimbabwe","Netherlands","Scotland","Nepal","Oman",
        "United Arab Emirates","United States","Namibia","Papua New Guinea"
    ]

    sel_ctr = st.selectbox(
        "Opponent",
        ["All"] + sorted(odi_nations),
        label_visibility="collapsed"
    )

with fc3:
    st.markdown('<div class="filter-label"> Select Year</div>', unsafe_allow_html=True)

    _yrs = (
        ["All"] + sorted([int(y) for y in _raw["Match_date"].dropna().dt.year.unique()])
        if not _raw.empty else ["All"]
    )

    sel_year = st.selectbox("Year", _yrs, label_visibility="collapsed")

st.markdown('<div class="filter-label">&nbsp;</div>', unsafe_allow_html=True)
analyze = st.button("ANALYZE NOW", use_container_width=True)

# Scenario detection
scenario = 1 if sel_ctr == "All" else 2
sc_desc  = ("Cross-country comparison charts" if scenario == 1
            else f"Venue-based analysis within {sel_ctr}")


st.divider()

# ── SESSION STATE ─────────────────────────────────────────────
if analyze:
    st.session_state.update(cw_ok=True, cw_player=sel_player,
                            cw_ctr=sel_ctr, cw_year=sel_year)

if not st.session_state.get("cw_ok"):
    st.markdown('<div class="no-data"><span class="icon"></span>'
                'Select a player, country &amp; year, then click <b>Analyze Now</b></div>',
                unsafe_allow_html=True)
    st.stop()

sel_player = st.session_state["cw_player"]
sel_ctr    = st.session_state["cw_ctr"]
sel_year   = st.session_state["cw_year"]
df_raw     = stats_dict.get(sel_player, pd.DataFrame())
scenario   = 1 if sel_ctr == "All" else 2

# ── PLAYER META ───────────────────────────────────────────────
prow = master_df[master_df["PlayerName"].str.strip() == sel_player]
if not prow.empty:
    pi        = prow.iloc[0]
    p_country = pi.get("Country",     "India")
    p_role    = pi.get("Role",        "Batsman")
    p_bat     = pi.get("BattingStyle","RHB")
    p_bowl    = pi.get("BowlingStyle","-")
    p_dob     = pi.get("DOB",         "N/A")
    p_debut   = pi.get("DebutYear",   "N/A")
    p_img     = pi.get("ImageFileName","")
else:
    p_country="India"; p_role="Batsman"; p_bat="RHB"
    p_bowl="-"; p_dob="N/A"; p_debut="N/A"; p_img=""

rl     = p_role.lower()
is_bat = "batsman"     in rl
is_bowl= "bowler"      in rl
is_all = "all-rounder" in rl or "all rounder" in rl
bcls   = ("badge-bowler" if is_bowl else ("badge-allrounder" if is_all else "badge-batsman"))

# ── BUILD WORKING DF ──────────────────────────────────────────
if df_raw.empty:
    st.markdown('<div class="no-data"><span class="icon"></span>No data for this player.</div>',
                unsafe_allow_html=True); st.stop()

wdf = df_raw.copy()
if sel_ctr != "All":
    wdf = wdf[wdf["Country"] == sel_ctr]
if sel_year != "All":
    wdf = wdf[wdf["Match_date"].dt.year == int(sel_year)]
wdf = wdf.sort_values("Match_date")

bat_df  = bat_rows(wdf)
bowl_df = bowl_rows(wdf)

if bat_df.empty and bowl_df.empty:
    st.markdown('<div class="no-data"><span class="icon"></span>No data for selected filters.</div>',
                unsafe_allow_html=True); st.stop()

total_m   = len(wdf)
ctr_label = sel_ctr if sel_ctr != "All" else "All Countries"
wins_m    = int((wdf["Match_result"]=="Won").sum())
losses_m  = int((wdf["Match_result"]=="Lose").sum())
win_pct   = round(wins_m / max(total_m,1) * 100, 1)

# ══════════════════════════════════════════════════════════════
#  SECTION 1 — PROFILE + COUNTRY INSIGHTS
# ══════════════════════════════════════════════════════════════
col_p, col_i = st.columns([1, 2.2], gap="large")

with col_p:
    st.markdown('<div class="sec-title"><span class="dot"></span>Player Profile</div>',
                unsafe_allow_html=True)
    img_src  = get_img_b64(p_img) if p_img else None
    img_html = (f'<img src="{img_src}" style="width:100px;height:100px;border-radius:50%;'
                f'object-fit:cover;border:3px solid #1B3A6B;display:block;margin:0 auto;">'
                if img_src else initials_avatar(sel_player))
    st.markdown(f"""
    <div class="profile-card">
      {img_html}
      <div class="player-name">{sel_player}</div>
      <div class="player-role-badge {bcls}">{p_role}</div>
      <div class="info-row"><span class="info-lbl">Batting</span>
        <span class="info-val">{p_bat}</span></div>
      <div class="info-row"><span class="info-lbl">Bowling</span>
        <span class="info-val">{p_bowl}</span></div>
      <div class="info-row"><span class="info-lbl">DOB</span>
        <span class="info-val">{p_dob}</span></div>
      <div class="info-row"><span class="info-lbl">Debut</span>
        <span class="info-val">{p_debut}</span></div>
      <div class="info-row"><span class="info-lbl">Scope</span>
        <span class="info-val">{ctr_label}</span></div>
    </div>""", unsafe_allow_html=True)

with col_i:
    nr_m = total_m - wins_m - losses_m
    st.markdown(f"""
    <div class="ctr-banner">
      <div>
        <div class="ctr-banner-lbl">
          {'Country Overview' if sel_ctr=='All' else 'Venue Analysis — ' + sel_ctr}
        </div>
        <div class="ctr-banner-val">{ctr_label}</div>
        <div class="ctr-banner-sub">
          {wins_m} Won &nbsp;·&nbsp; {losses_m} Lost &nbsp;·&nbsp;
          {nr_m} N/R &nbsp;·&nbsp; {total_m} Matches
        </div>
      </div>
      <div style="text-align:right;">
        <div class="ctr-banner-pct">{win_pct}%</div>
        <div class="ctr-banner-pct-lbl">Win Rate</div>
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown('<div class="sec-title"><span class="dot"></span>Country Insights</div>',
                unsafe_allow_html=True)

    # ── compute insights from FULL df_raw ─────────────────────
    full_bat  = bat_rows(df_raw)
    full_bowl = bowl_rows(df_raw)
    ins_html  = ""

    if is_bat and not is_all:
        if not full_bat.empty and "Country" in full_bat:
            # Best country by avg
            ca = (full_bat.groupby("Country")
                  .apply(lambda g: calc_avg(g["Runs"],
                         g.get("Not_out",pd.Series(0,index=g.index)), len(g)))
                  .reset_index(name="Avg"))
            if len(ca) >= 1:
                bc = ca.loc[ca["Avg"].idxmax()]
                ins_html += ins("", f"Best country: <strong>{bc['Country']}</strong> (Avg <strong>{bc['Avg']}</strong>)")
            # SENA vs Asia avg
            sena_b = full_bat[full_bat["Country"].isin(SENA)]
            asia_b = full_bat[full_bat["Country"].isin(ASIA)]
            if not sena_b.empty:
                sa = calc_avg(sena_b["Runs"], sena_b.get("Not_out",pd.Series(0,index=sena_b.index)), len(sena_b))
                ins_html += ins("", f"Average in <strong>SENA</strong> countries: <strong>{sa}</strong>")
            if not asia_b.empty:
                aa = calc_avg(asia_b["Runs"], asia_b.get("Not_out",pd.Series(0,index=asia_b.index)), len(asia_b))
                ins_html += ins("", f"Average in <strong>Asia</strong> countries: <strong>{aa}</strong>")
            # Best venue
            if "Venue" in full_bat.columns:
                va = (full_bat.groupby("Venue")
                      .apply(lambda g: calc_avg(g["Runs"],
                             g.get("Not_out",pd.Series(0,index=g.index)), len(g)))
                      .reset_index(name="Avg"))
                if not va.empty:
                    bv = va.loc[va["Avg"].idxmax()]
                    ins_html += ins("", f"Best venue: <strong>{bv['Venue']}</strong> (Avg <strong>{bv['Avg']}</strong>)")
            # Home vs Away
            home_b = full_bat[full_bat["Country"]=="India"]
            away_b = full_bat[full_bat["Country"]!="India"]
            if not home_b.empty and not away_b.empty:
                h_avg = calc_avg(home_b["Runs"],home_b.get("Not_out",pd.Series(0,index=home_b.index)),len(home_b))
                a_avg = calc_avg(away_b["Runs"],away_b.get("Not_out",pd.Series(0,index=away_b.index)),len(away_b))
                better = "home" if h_avg >= a_avg else "away"
                ins_html += ins("", f"Home avg <strong>{h_avg}</strong> vs Away avg <strong>{a_avg}</strong> — performs better <strong>{better}</strong>")

    elif is_bowl and not is_all:
        if not full_bowl.empty and "Country" in full_bowl:
            cw = full_bowl.groupby("Country")["Wickets"].sum()
            if not cw.empty:
                bc2 = cw.idxmax()
                ins_html += ins("", f"Most wickets in <strong>{bc2}</strong> ({int(cw[bc2])} wickets)")
            ce = full_bowl.groupby("Country")["Economy"].mean().round(2)
            if not ce.empty:
                bce = ce.idxmin()
                ins_html += ins("", f"Best economy in <strong>{bce}</strong> ({ce[bce]:.2f} RPO)")
            # SENA vs Asia wickets
            sena_bw = full_bowl[full_bowl["Country"].isin(SENA)]
            asia_bw = full_bowl[full_bowl["Country"].isin(ASIA)]
            if not sena_bw.empty:
                ins_html += ins("", f"Wickets in <strong>SENA</strong>: <strong>{int(sena_bw['Wickets'].sum())}</strong>")
            if not asia_bw.empty:
                ins_html += ins("", f"Wickets in <strong>Asia</strong>: <strong>{int(asia_bw['Wickets'].sum())}</strong>")
            if "Venue" in full_bowl.columns:
                vw = full_bowl.groupby("Venue")["Wickets"].sum()
                if not vw.empty:
                    bvw = vw.idxmax()
                    ins_html += ins("", f"Best venue: <strong>{bvw}</strong> ({int(vw[bvw])} wickets)")

    else:  # All-rounder
        if not full_bat.empty and "Country" in full_bat:
            ca3 = (full_bat.groupby("Country")
                   .apply(lambda g: calc_avg(g["Runs"],
                          g.get("Not_out",pd.Series(0,index=g.index)), len(g)))
                   .reset_index(name="Avg"))
            if not ca3.empty:
                bc3 = ca3.loc[ca3["Avg"].idxmax()]
                ins_html += ins("", f"Best batting country: <strong>{bc3['Country']}</strong> (Avg <strong>{bc3['Avg']}</strong>)")
        if not full_bowl.empty and "Country" in full_bowl:
            cw3 = full_bowl.groupby("Country")["Wickets"].sum()
            if not cw3.empty:
                bc4 = cw3.idxmax()
                ins_html += ins("", f"Most wickets in <strong>{bc4}</strong> ({int(cw3[bc4])} wickets)")
        if not bat_df.empty and not bowl_df.empty:
            tr3 = int(bat_df["Runs"].sum()); tw3 = int(bowl_df["Wickets"].sum())
            ins_html += ins("", f"All-round in {ctr_label}: <strong>{tr3}</strong> runs + <strong>{tw3}</strong> wickets")
        sena_b3 = bat_rows(df_raw[df_raw["Country"].isin(SENA)]) if "Country" in df_raw else pd.DataFrame()
        asia_b3 = bat_rows(df_raw[df_raw["Country"].isin(ASIA)]) if "Country" in df_raw else pd.DataFrame()
        if not sena_b3.empty:
            sa3 = calc_avg(sena_b3["Runs"],sena_b3.get("Not_out",pd.Series(0,index=sena_b3.index)),len(sena_b3))
            ins_html += ins("", f"SENA batting avg: <strong>{sa3}</strong>")
        if not asia_b3.empty:
            aa3 = calc_avg(asia_b3["Runs"],asia_b3.get("Not_out",pd.Series(0,index=asia_b3.index)),len(asia_b3))
            ins_html += ins("", f"Asia batting avg: <strong>{aa3}</strong>")

    st.markdown(ins_html or
        '<div class="no-data" style="padding:20px">No insights available.</div>',
        unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SECTION 2 — KPI CARDS
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="sec-title"><span class="dot"></span>Key Performance Indicators</div>',
            unsafe_allow_html=True)

if is_bat and not is_all:
    inn   = len(bat_df)
    tot_r = int(bat_df["Runs"].sum())
    no_s  = bat_df.get("Not_out", pd.Series(0, index=bat_df.index))
    avg   = calc_avg(bat_df["Runs"], no_s, inn)
    avsr  = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(), 1)
    f4    = int(bat_df["4s"].sum()); f6 = int(bat_df["6s"].sum())
    fif   = int(bat_df["50s"].sum()) if "50s" in bat_df else 0
    hun   = int(bat_df["100s"].sum()) if "100s" in bat_df else 0
    hs    = (f"{int(bat_df.loc[bat_df['Runs'].idxmax(),'Runs'])}"
             f"{'*' if int(bat_df.loc[bat_df['Runs'].idxmax(),'Not_out']) else ''}"
             if not bat_df.empty else "0")
    kpis  = [
        (str(inn),        "Innings",    f"in {ctr_label}", ""),
        (str(tot_r),      "Total Runs", "Aggregate",       "hl"),
        (str(avg),        "Average",    "Runs/dismissal",  ""),
        (str(avsr),       "Strike Rate","Avg SR",          ""),
        (hs,              "Best Score", "Highest inn",     "hl"),
        (f"{f4}/{f6}",    "4s / 6s",   "Boundaries",      ""),
        (f"{fif}/{hun}",  "50s/100s",  "Milestones",      "hl-g"),
    ]

elif is_bowl and not is_all:
    inn    = len(bowl_df)
    tballs = int(bowl_df["Balls_Bowled"].sum()) if "Balls_Bowled" in bowl_df else 0
    ovs    = f"{tballs//6}.{tballs%6}"
    twkts  = int(bowl_df["Wickets"].sum()) if "Wickets" in bowl_df else 0
    rc     = int(bowl_df["Runs_Conceded"].sum()) if "Runs_Conceded" in bowl_df else 0
    ba     = round(rc/max(twkts,1), 2)
    eco    = round(bowl_df["Economy"].replace(0,float("nan")).mean(), 2) if "Economy" in bowl_df else 0
    bsr_v  = round(tballs/max(twkts,1), 1)
    dots   = int(get_col(bowl_df,"Dot_Balls").sum())
    if twkts > 0:
        bfr = bowl_df.sort_values(["Wickets","Runs_Conceded"],ascending=[False,True]).iloc[0]
        bfig= f"{int(bfr['Wickets'])}/{int(bfr.get('Runs_Conceded',0))}"
    else: bfig="0/0"
    kpis = [
        (str(inn),  "Matches",     f"in {ctr_label}",  ""),
        (ovs,       "Overs",       "Total bowled",      ""),
        (str(twkts),"Wickets",     "Total",             "hl"),
        (str(eco),  "Economy",     "Runs/over",         ""),
        (str(ba),   "Bowl Average","Runs/wicket",       ""),
        (str(bsr_v),"Bowl SR",     "Balls/wicket",      ""),
        (bfig,      "Best Figures","Wkts/Runs",         "hl"),
    ]

else:  # all-rounder
    bi   = len(bat_df)
    tr2  = int(bat_df["Runs"].sum()) if not bat_df.empty else 0
    no2  = bat_df.get("Not_out",pd.Series(0,index=bat_df.index)) if not bat_df.empty else pd.Series([0])
    av2  = calc_avg(bat_df["Runs"] if not bat_df.empty else pd.Series([0]), no2, max(bi,1))
    sr2  = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(),1) if not bat_df.empty else 0
    fi2  = int(bat_df["50s"].sum()) if not bat_df.empty and "50s" in bat_df else 0
    hu2  = int(bat_df["100s"].sum()) if not bat_df.empty and "100s" in bat_df else 0
    tw2  = int(bowl_df["Wickets"].sum()) if not bowl_df.empty else 0
    ec2  = round(bowl_df["Economy"].replace(0,float("nan")).mean(),2) if not bowl_df.empty and "Economy" in bowl_df else 0
    tb2  = int(bowl_df["Balls_Bowled"].sum()) if not bowl_df.empty and "Balls_Bowled" in bowl_df else 0
    ov2  = f"{tb2//6}.{tb2%6}"
    kpis = [
        (str(bi),       "Innings",    f"in {ctr_label}", ""),
        (str(tr2),      "Total Runs", "Batting",         "hl"),
        (str(av2),      "Bat Average","Runs/dismissal",  ""),
        (str(sr2),      "Strike Rate","Avg SR",          ""),
        (f"{fi2}/{hu2}","50s/100s",  "Milestones",       "hl-g"),
        (str(tw2),      "Wickets",   "Bowling",          "hl"),
        (str(ec2),      "Economy",   "Runs/over",        ""),
        (ov2,           "Overs",     "Bowled",           ""),
    ]

kcols = st.columns(len(kpis))
for i,(v,l,s,var) in enumerate(kpis):
    with kcols[i]:
        st.markdown(kcard(v,l,s,var), unsafe_allow_html=True)
st.divider()

# ── AGGREGATION HELPERS ────────────────────────────────────────
def agg_bat_grp(df, grp="Country"):
    out = []
    for key, g in df.groupby(grp):
        inn = len(g); tr = int(g["Runs"].sum())
        no  = g.get("Not_out", pd.Series(0,index=g.index))
        avg_= calc_avg(g["Runs"], no, inn)
        sr_ = round(g["Strike_Rate"].replace(0,float("nan")).mean(), 1)
        f_  = int(g["4s"].sum()); s_ = int(g["6s"].sum())
        d_  = int(g["Dots"].sum()); sg_= int(g["1s_2s_3s"].sum()) if "1s_2s_3s" in g else 0
        fi_ = int(g["50s"].sum()) if "50s" in g else 0
        hu_ = int(g["100s"].sum()) if "100s" in g else 0
        bnd_r = f_*4 + s_*6
        out.append({grp:key, "Innings":inn, "Runs":tr, "Average":avg_,
                    "SR":sr_, "Fours":f_, "Sixes":s_, "Dots":d_, "Singles":sg_,
                    "Fifties":fi_, "Hundreds":hu_,
                    "Bnd_Pct":round(bnd_r/max(tr,1)*100,1)})
    return pd.DataFrame(out)

def agg_bowl_grp(df, grp="Country"):
    out = []
    for key, g in df.groupby(grp):
        wk = int(g["Wickets"].sum())
        rc = int(g["Runs_Conceded"].sum()) if "Runs_Conceded" in g else 0
        ec = round(g["Economy"].replace(0,float("nan")).mean(),2) if "Economy" in g else 0
        bl = int(g["Balls_Bowled"].sum()) if "Balls_Bowled" in g else 0
        bsr_= round(bl/max(wk,1),1)
        ba_ = round(rc/max(wk,1),2)
        fd  = int(g["Full_Dismissals"].sum()) if "Full_Dismissals" in g else 0
        gd  = int(get_col(g,"Good_Dissmisals","Good_Dismissals").sum())
        sd  = int(get_col(g,"Short_Dissmisals","Short_Dismissals").sum())
        rw  = int(g["Wicket_RHB"].sum()) if "Wicket_RHB" in g else 0
        lw  = int(g["Wicket_LHB"].sum()) if "Wicket_LHB" in g else 0
        db  = int(get_col(g,"Dot_Balls").sum())
        out.append({grp:key, "Wickets":wk, "Runs_C":rc, "Economy":ec,
                    "Bowl_SR":bsr_, "Bowl_Avg":ba_, "Full":fd, "Good":gd,
                    "Short":sd, "RHB_W":rw, "LHB_W":lw, "Dots":db})
    return pd.DataFrame(out)

# ── base data for scenario ─────────────────────────────────────
base = df_raw.copy()
if sel_year != "All":
    base = base[base["Match_date"].dt.year == int(sel_year)]

def cch(name):
    return cclr(name)

# ══════════════════════════════════════════════════════════════
#  SCENARIO 1 — CROSS-COUNTRY CHARTS
# ══════════════════════════════════════════════════════════════
if scenario == 1:

    # ── BATSMAN ──────────────────────────────────────────────
    if is_bat and not is_all:
        ogg = agg_bat_grp(bat_rows(base))
        if ogg.empty:
            st.markdown('<div class="no-data">No batting data available.</div>',unsafe_allow_html=True)
        else:
            clrs = cclrs(ogg["Country"])

            # Row 1 — 3 charts
            r1a, r1b, r1c = st.columns(3, gap="large")

            # 1. Runs by Country — Treemap
            with r1a:
                
                fig = px.treemap(ogg, path=["Country"], values="Runs",
                    color="Runs",
                    color_continuous_scale=[[0,"#EBF3FF"],[0.5,STEEL],[1,NAVY]],
                    hover_data={"Average":True,"Innings":True})
                fig.update_traces(textinfo="label+value",
                    textfont=dict(family="Poppins",size=13))
                fig.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", height=310,
                    margin=dict(l=10,r=10,t=44,b=10),
                    title=dict(text="Runs by Country",
                        font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
                    coloraxis_showscale=False)
                pch(fig)
                

            # 2. Batting Average by Country — Box Plot (per-match values)
            with r1b:
                
                bp = bat_rows(base)[["Country","Runs"]].dropna()
                fig2 = go.Figure()
                for i,(ctr,grp) in enumerate(bp.groupby("Country")):
                    fig2.add_trace(go.Box(
                        y=grp["Runs"], name=ctr, boxpoints="all",
                        marker=dict(color=cclr(ctr), size=5,
                            line=dict(color="white",width=1)),
                        line=dict(color=cclr(ctr), width=1.5),
                        fillcolor=cclr(ctr),
                        opacity=0.3
                    ))
                pchart(fig2, 310, "Runs Distribution by Country")
                fig2.update_yaxes(title_text="Runs per Innings")
                pch(fig2)
                

            # 3. Strike Rate by Country — Radar Chart
            with r1c:
                
                cats   = list(ogg["Country"]) + [ogg["Country"].iloc[0]]
                sr_vals= list(ogg["SR"]) + [ogg["SR"].iloc[0]]
                avg_vals=list(ogg["Average"]) + [ogg["Average"].iloc[0]]
                fig3 = go.Figure()
                fig3.add_trace(go.Scatterpolar(
                    r=sr_vals, theta=cats, fill="toself",
                    name="Strike Rate",
                    line=dict(color=NAVY, width=2),
                    fillcolor=NAVY,
                    opacity=0.2))
                fig3.add_trace(go.Scatterpolar(
                    r=avg_vals, theta=cats, fill="toself",
                    name="Average",
                    line=dict(color=AMBER, width=2, dash="dot"),
                    fillcolor=AMBER,
                    opacity=0.2))
                fig3.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", height=310,
                    margin=dict(l=30,r=30,t=44,b=10),
                    title=dict(text="Strike Rate & Average by Country (Radar)",
                        font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
                    polar=dict(
                        radialaxis=dict(visible=True,gridcolor="#f0f3f8",
                            linecolor=BORDER,tickfont=dict(size=8)),
                        angularaxis=dict(tickfont=dict(family="Inter",size=9))),
                    legend=dict(bgcolor="rgba(255,255,255,.9)",bordercolor=BORDER,
                        borderwidth=1,font=dict(size=10)))
                pch(fig3)
                

            # Row 2 — 3 charts
            r2a, r2b, r2c = st.columns(3, gap="large")

            # 4. Boundary % — 100% Stacked Bar
            with r2a:
                
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(name="4s Runs", x=ogg["Country"],
                    y=ogg["Fours"]*4, marker_color=NAVY))
                fig4.add_trace(go.Bar(name="6s Runs", x=ogg["Country"],
                    y=ogg["Sixes"]*6, marker_color=AMBER))
                fig4.add_trace(go.Bar(name="1s/2s/3s", x=ogg["Country"],
                    y=ogg["Singles"], marker_color=STEEL))
                fig4.add_trace(go.Bar(name="Dots", x=ogg["Country"],
                    y=ogg["Dots"], marker_color="#D5DDE5"))
                fig4.update_layout(barmode="stack", barnorm="percent")
                pchart(fig4, 310, "Boundary % Composition by Country")
                fig4.update_xaxes(tickangle=-20)
                fig4.update_yaxes(title_text="%")
                pch(fig4)
                

            # 5. Dismissal Length by Country — Heatmap
            with r2b:
                
                b_all = bat_rows(base)
                dl    = b_all[~b_all["Dismissed_Length"].astype(str).isin(
                    ["Not Applicable","0","nan",""])]
                if not dl.empty:
                    heat = dl.groupby(["Country","Dismissed_Length"]).size().unstack(fill_value=0)
                    fig5 = go.Figure(go.Heatmap(
                        z=heat.values, x=heat.columns.tolist(), y=heat.index.tolist(),
                        colorscale=[[0,"#EBF3FF"],[0.5,"#6B9FCC"],[1,NAVY]],
                        text=heat.values, texttemplate="%{text}",
                        textfont=dict(family="Poppins",size=12),
                        showscale=True, colorbar=dict(thickness=12,len=0.8)))
                    pchart(fig5, 310, "Dismissal Length by Country (Heatmap)")
                    fig5.update_xaxes(title_text="Dismissal Length",tickangle=-15)
                    fig5.update_yaxes(title_text="Country")
                    pch(fig5)
                else:
                    st.markdown('<div class="no-data">No dismissal length data.</div>',unsafe_allow_html=True)
                

            # 6. SENA vs Asia — Grouped Bar
            with r2c:
                
                sena_d = bat_rows(base[base["Country"].isin(SENA)])
                asia_d = bat_rows(base[base["Country"].isin(ASIA)])
                if not sena_d.empty or not asia_d.empty:
                    def grp_stats(df, label):
                        inn_ = len(df)
                        tr_  = int(df["Runs"].sum())
                        no_  = df.get("Not_out",pd.Series(0,index=df.index))
                        av_  = calc_avg(df["Runs"],no_,inn_) if inn_ else 0
                        sr_  = round(df["Strike_Rate"].replace(0,float("nan")).mean(),1) if inn_ else 0
                        return {"label":label,"Avg":av_,"SR":sr_,"Runs":tr_}
                    s_ = grp_stats(sena_d, "SENA")
                    a_ = grp_stats(asia_d, "Asia")
                    cats_ = ["Average","Strike Rate","Runs"]
                    s_vals= [s_["Avg"], s_["SR"], s_["Runs"]]
                    a_vals= [a_["Avg"], a_["SR"], a_["Runs"]]
                    # normalise for radar-style grouped bar
                    fig6 = go.Figure()
                    fig6.add_trace(go.Bar(name="SENA", x=cats_,
                        y=s_vals, marker_color=RED))
                    fig6.add_trace(go.Bar(name="Asia", x=cats_,
                        y=a_vals, marker_color=GREEN))
                    fig6.update_layout(barmode="group")
                    pchart(fig6, 310, "SENA vs Asia Performance")
                    pch(fig6)
                else:
                    st.markdown('<div class="no-data">No SENA/Asia data.</div>',unsafe_allow_html=True)
                

    # ── BOWLER S1 ─────────────────────────────────────────────
    elif is_bowl and not is_all:
        bgg = agg_bowl_grp(bowl_rows(base))
        if bgg.empty:
            st.markdown('<div class="no-data">No bowling data available.</div>',unsafe_allow_html=True)
        else:
            clrs = cclrs(bgg["Country"])

            # Row 1 — 3 charts
            r1a, r1b, r1c = st.columns(3, gap="large")

            # 1. Wickets by Country — Treemap
            with r1a:
                
                fig = px.treemap(bgg, path=["Country"], values="Wickets",
                    color="Wickets",
                    color_continuous_scale=[[0,"#EBF3FF"],[0.5,STEEL],[1,NAVY]])
                fig.update_traces(textinfo="label+value",
                    textfont=dict(family="Poppins",size=13))
                fig.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", height=310,
                    margin=dict(l=10,r=10,t=44,b=10),
                    title=dict(text="Wickets by Country",
                        font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
                    coloraxis_showscale=False)
                pch(fig)
                

            # 2. Economy by Country — Violin (using per-match data)
            with r1b:
                
                bwf = bowl_rows(base)[["Country","Economy"]].dropna()
                bwf = bwf[bwf["Economy"] > 0]
                fig2 = go.Figure()
                for ctr, grp in bwf.groupby("Country"):
                    fig2.add_trace(go.Violin(
                        y=grp["Economy"], name=ctr, box_visible=True,
                        meanline_visible=True, points="all",
                        line=dict(color=cclr(ctr), width=1.5),
                        fillcolor=cclr(ctr),
                        marker=dict(color=cclr(ctr), size=5),
                    ))
                fig2.add_hline(y=7, line_dash="dash", line_color=AMBER,
                    annotation_text="Economy 7",
                    annotation_font=dict(color=AMBER,size=9))
                pchart(fig2, 310, "Economy Distribution by Country")
                fig2.update_yaxes(title_text="Economy (RPO)")
                pch(fig2)
                

            # 3. Bowling Control — Bubble Chart (Wickets vs Economy vs Country)
            with r1c:
                
                fig3 = go.Figure(go.Scatter(
                    x=bgg["Economy"], y=bgg["Wickets"],
                    mode="markers+text",
                    text=bgg["Country"],
                    textposition="top center",
                    textfont=dict(family="Poppins",size=9,color=NAVY),
                    marker=dict(
                        size=[max(w*8, 14) for w in bgg["Wickets"]],
                        color=clrs,
                        line=dict(color="white",width=2),
                        opacity=0.85),
                ))
                pchart(fig3, 310, "Bowling Control — Wickets vs Economy (Bubble)")
                fig3.update_xaxes(title_text="Economy (RPO)")
                fig3.update_yaxes(title_text="Wickets")
                pch(fig3)
                

            # Row 2 — 2 charts
            r2a, r2b = st.columns(2, gap="large")

            # 4. Wickets by Length (by Country) — Stacked Area
            with r2a:
                
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(name="Good Length", x=bgg["Country"],
                    y=bgg["Good"], marker_color=NAVY))
                fig4.add_trace(go.Bar(name="Full Length", x=bgg["Country"],
                    y=bgg["Full"], marker_color=AMBER))
                fig4.add_trace(go.Bar(name="Short", x=bgg["Country"],
                    y=bgg["Short"], marker_color=STEEL))
                fig4.update_layout(barmode="stack")
                pchart(fig4, 310, "Wickets by Pitch Length per Country")
                fig4.update_xaxes(tickangle=-20)
                fig4.update_yaxes(title_text="Wickets")
                pch(fig4)
                

            # 5. Bowling SR by Country — Column Chart
            with r2b:
                
                bsr_clr = [GREEN if b<=20 else (AMBER if b<=30 else RED)
                           for b in bgg["Bowl_SR"]]
                fig5 = go.Figure(go.Bar(
                    x=bgg["Country"], y=bgg["Bowl_SR"],
                    marker=dict(color=bsr_clr, line=dict(color="white",width=1)),
                    text=bgg["Bowl_SR"], textposition="outside",
                    textfont=dict(family="Poppins",size=11),
                ))
                fig5.add_hline(y=25, line_dash="dash", line_color=AMBER,
                    annotation_text="BSR 25",
                    annotation_font=dict(color=AMBER,size=9))
                pchart(fig5, 310, "Bowling Strike Rate by Country")
                fig5.update_xaxes(tickangle=-20)
                fig5.update_yaxes(title_text="Balls per Wicket")
                pch(fig5)
                

    # ── ALL-ROUNDER S1 ────────────────────────────────────────
    else:
        ogg = agg_bat_grp(bat_rows(base))
        bgg = agg_bowl_grp(bowl_rows(base))

        r1a, r1b = st.columns(2, gap="large")

        # 1. Dual Axis: Runs (bar) + Wickets (line) by Country
        with r1a:
            
            if not ogg.empty and not bgg.empty:
                mg = pd.merge(ogg[["Country","Runs","Average"]],
                              bgg[["Country","Wickets","Economy"]],
                              on="Country", how="outer").fillna(0)
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Runs", x=mg["Country"], y=mg["Runs"],
                    marker=dict(color=cclrs(mg["Country"]),
                        line=dict(color="white",width=1)),
                    text=mg["Runs"], textposition="outside",
                    textfont=dict(family="Poppins",size=10)))
                fig.add_trace(go.Scatter(name="Wickets", x=mg["Country"], y=mg["Wickets"],
                    mode="lines+markers+text",
                    line=dict(color=AMBER, width=2.5),
                    marker=dict(size=10,color=AMBER,line=dict(color="white",width=2)),
                    text=mg["Wickets"], textposition="top center",
                    textfont=dict(family="Poppins",size=10,color=AMBER),
                    yaxis="y2"))
                fig.update_layout(
                    yaxis2=dict(overlaying="y", side="right",
                        gridcolor="#f0f3f8", linecolor=BORDER,
                        tickfont=dict(size=9), title_text="Wickets"))
                pchart(fig, 310, "Bat vs Ball Performance by Country (Dual Axis)")
                fig.update_xaxes(tickangle=-20)
                fig.update_yaxes(title_text="Runs")
                pch(fig)
            

        # 2. Batting Average by Country — bar
        with r1b:
            
            if not ogg.empty:
                fig2 = go.Figure(go.Bar(
                    x=ogg["Country"], y=ogg["Average"],
                    marker=dict(color=cclrs(ogg["Country"]),
                        line=dict(color="white",width=1)),
                    text=ogg["Average"], textposition="outside",
                    textfont=dict(family="Poppins",size=11),
                ))
                fig2.add_hline(y=50, line_dash="dash", line_color=AMBER,
                    annotation_text="Avg 50",
                    annotation_font=dict(color=AMBER,size=9))
                pchart(fig2, 310, "Batting Average by Country")
                fig2.update_xaxes(tickangle=-20)
                fig2.update_yaxes(title_text="Average")
                pch(fig2)
            

        r2a, r2b = st.columns(2, gap="large")

        # 3. Wickets treemap
        with r2a:
            
            if not bgg.empty:
                fig3 = px.treemap(bgg, path=["Country"], values="Wickets",
                    color="Wickets",
                    color_continuous_scale=[[0,"#EBF3FF"],[0.5,STEEL],[1,NAVY]])
                fig3.update_traces(textinfo="label+value",
                    textfont=dict(family="Poppins",size=13))
                fig3.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", height=310,
                    margin=dict(l=10,r=10,t=44,b=10),
                    title=dict(text="Wickets by Country (Treemap)",
                        font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
                    coloraxis_showscale=False)
                pch(fig3)
            

        # 4. Boundary % stacked
        with r2b:
            st.markdown("", unsafe_allow_html=True)
            st.container(border=True)
            if not ogg.empty:
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(name="4s Runs", x=ogg["Country"],
                    y=ogg["Fours"]*4, marker_color=NAVY))
                fig4.add_trace(go.Bar(name="6s Runs", x=ogg["Country"],
                    y=ogg["Sixes"]*6, marker_color=AMBER))
                fig4.add_trace(go.Bar(name="1s/2s/3s", x=ogg["Country"],
                    y=ogg["Singles"], marker_color=STEEL))
                fig4.add_trace(go.Bar(name="Dots", x=ogg["Country"],
                    y=ogg["Dots"], marker_color="#D5DDE5"))
                fig4.update_layout(barmode="stack", barnorm="percent")
                pchart(fig4, 310, "Boundary % Composition by Country")
                fig4.update_xaxes(tickangle=-20)
                pch(fig4)
            

# ══════════════════════════════════════════════════════════════
#  SCENARIO 2 — VENUE-BASED ANALYSIS (specific country)
# ══════════════════════════════════════════════════════════════
else:
    # Use full year range unless year is filtered
    base2 = df_raw[df_raw["Country"] == sel_ctr].copy()
    if sel_year != "All":
        base2 = base2[base2["Match_date"].dt.year == int(sel_year)]

    # ── BATSMAN S2 ────────────────────────────────────────────
    if is_bat and not is_all:
        b2  = bat_rows(base2)
        ogg2= agg_bat_grp(b2, "Venue") if not b2.empty else pd.DataFrame()
        if ogg2.empty:
            st.markdown('<div class="no-data">No venue data.</div>',unsafe_allow_html=True)
        else:
            vclrs = [CHART_COLORS[i % len(CHART_COLORS)] for i in range(len(ogg2))]

            r1a, r1b, r1c = st.columns(3, gap="large")

            # 1. Runs by Venue — Treemap
            with r1a:
                fig = px.treemap(ogg2, path=["Venue"], values="Runs",
                    color="Runs",
                    color_continuous_scale=[[0,"#EBF3FF"],[0.5,STEEL],[1,NAVY]])
                fig.update_traces(textinfo="label+value",
                    textfont=dict(family="Poppins",size=13))
                fig.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", height=310,
                    margin=dict(l=10,r=10,t=44,b=10),
                    title=dict(text=f"Runs by Venue in {sel_ctr}",
                        font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
                    coloraxis_showscale=False)
                pch(fig)
                

            # 2. Average by Venue — Bar
            with r1b:
                
                avg_clr2 = [GREEN if a>=50 else (AMBER if a>=30 else RED) for a in ogg2["Average"]]
                fig2 = go.Figure(go.Bar(
                    y=ogg2["Venue"], x=ogg2["Average"], orientation="h",
                    marker=dict(color=avg_clr2, line=dict(color="white",width=1)),
                    text=ogg2["Average"], textposition="outside",
                    textfont=dict(family="Poppins",size=11),
                ))
                fig2.add_vline(x=50, line_dash="dash", line_color=AMBER,
                    annotation_text="Avg 50",
                    annotation_font=dict(color=AMBER,size=9))
                pchart(fig2, 310, f"Batting Average by Venue in {sel_ctr}")
                fig2.update_xaxes(title_text="Average")
                pch(fig2)
                

            # 3. Strike Rate by Venue — Dot Plot
            with r1c:
                
                sr2_clr = [GREEN if s>=100 else (AMBER if s>=70 else RED) for s in ogg2["SR"]]
                fig3 = go.Figure(go.Scatter(
                    x=ogg2["Venue"], y=ogg2["SR"],
                    mode="markers+text",
                    text=[str(s) for s in ogg2["SR"]],
                    textposition="top center",
                    textfont=dict(family="Poppins",size=9,color=NAVY),
                    marker=dict(size=16, color=sr2_clr,
                        line=dict(color="white",width=2)),
                ))
                fig3.add_hline(y=100, line_dash="dash", line_color=GREEN,
                    annotation_text="SR 100",
                    annotation_font=dict(color=GREEN,size=9))
                pchart(fig3, 310, f"Strike Rate by Venue in {sel_ctr}")
                fig3.update_xaxes(tickangle=-20)
                fig3.update_yaxes(title_text="Strike Rate")
                pch(fig3)
                

            r2a, r2b, r2c = st.columns(3, gap="large")

            # 4. Boundary % by Venue — 100% Stacked Bar
            with r2a:
                
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(name="4s Runs", x=ogg2["Venue"],
                    y=ogg2["Fours"]*4, marker_color=NAVY))
                fig4.add_trace(go.Bar(name="6s Runs", x=ogg2["Venue"],
                    y=ogg2["Sixes"]*6, marker_color=AMBER))
                fig4.add_trace(go.Bar(name="1s/2s/3s", x=ogg2["Venue"],
                    y=ogg2["Singles"], marker_color=STEEL))
                fig4.add_trace(go.Bar(name="Dots", x=ogg2["Venue"],
                    y=ogg2["Dots"], marker_color="#D5DDE5"))
                fig4.update_layout(barmode="stack", barnorm="percent")
                pchart(fig4, 310, f"Boundary % by Venue in {sel_ctr}")
                fig4.update_xaxes(tickangle=-20)
                fig4.update_yaxes(title_text="%")
                pch(fig4)
                

            # 5. Dismissal Length by Venue — Heatmap
            with r2b:
                
                dl2 = b2[~b2["Dismissed_Length"].astype(str).isin(
                    ["Not Applicable","0","nan",""])]
                if not dl2.empty:
                    heat2 = dl2.groupby(["Venue","Dismissed_Length"]).size().unstack(fill_value=0)
                    fig5  = go.Figure(go.Heatmap(
                        z=heat2.values, x=heat2.columns.tolist(), y=heat2.index.tolist(),
                        colorscale=[[0,"#EBF3FF"],[0.5,"#6B9FCC"],[1,NAVY]],
                        text=heat2.values, texttemplate="%{text}",
                        textfont=dict(family="Poppins",size=12),
                        showscale=True, colorbar=dict(thickness=12,len=0.8)))
                    pchart(fig5, 310, f"Dismissal Length by Venue in {sel_ctr}")
                    fig5.update_xaxes(tickangle=-15)
                    fig5.update_yaxes(title_text="Venue")
                    pch(fig5)
                else:
                    st.markdown('<div class="no-data">No dismissal data.</div>',unsafe_allow_html=True)
                

            # 6. Runs vs Innings per venue — Scatter bubble
            with r2c:
                
                fig6 = go.Figure(go.Scatter(
                    x=ogg2["Innings"], y=ogg2["Runs"],
                    mode="markers+text",
                    text=ogg2["Venue"],
                    textposition="top center",
                    textfont=dict(family="Poppins",size=9,color=NAVY),
                    marker=dict(
                        size=[max(a/5,10) for a in ogg2["Average"]],
                        color=vclrs,
                        line=dict(color="white",width=2),
                        opacity=0.85),
                    hovertemplate="<b>%{text}</b><br>Innings: %{x}<br>Runs: %{y}<extra></extra>",
                ))
                pchart(fig6, 310, f"Runs vs Innings by Venue — {sel_ctr} (size=avg)")
                fig6.update_xaxes(title_text="Innings")
                fig6.update_yaxes(title_text="Runs")
                pch(fig6)
                

    # ── BOWLER S2 ─────────────────────────────────────────────
    elif is_bowl and not is_all:
        bw2  = bowl_rows(base2)
        bgg2 = agg_bowl_grp(bw2, "Venue") if not bw2.empty else pd.DataFrame()
        if bgg2.empty:
            st.markdown('<div class="no-data">No venue bowling data.</div>',unsafe_allow_html=True)
        else:
            vclrs2 = [CHART_COLORS[i % len(CHART_COLORS)] for i in range(len(bgg2))]

            r1a, r1b, r1c = st.columns(3, gap="large")

            # 1. Wickets by Venue — Treemap
            with r1a:
                
                fig = px.treemap(bgg2, path=["Venue"], values="Wickets",
                    color="Wickets",
                    color_continuous_scale=[[0,"#EBF3FF"],[0.5,STEEL],[1,NAVY]])
                fig.update_traces(textinfo="label+value",
                    textfont=dict(family="Poppins",size=13))
                fig.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)", height=310,
                    margin=dict(l=10,r=10,t=44,b=10),
                    title=dict(text=f"Wickets by Venue in {sel_ctr}",
                        font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
                    coloraxis_showscale=False)
                pch(fig)
                

            # 2. Economy by Venue — Box/Violin
            with r1b:
                
                bvec = bw2[["Venue","Economy"]].dropna()
                bvec = bvec[bvec["Economy"]>0]
                fig2 = go.Figure()
                for i,(ven,grp) in enumerate(bvec.groupby("Venue")):
                    c_ = CHART_COLORS[i % len(CHART_COLORS)]
                    fig2.add_trace(go.Box(
                        y=grp["Economy"], name=ven, boxpoints="all",
                        marker=dict(color=c_,size=5,line=dict(color="white",width=1)),
                        line=dict(color=cclr(grp), width=1.5),
                        fillcolor=cclr(grp),
                    ))
                fig2.add_hline(y=7, line_dash="dash", line_color=AMBER,
                    annotation_text="Economy 7",
                    annotation_font=dict(color=AMBER,size=9))
                pchart(fig2, 310, f"Economy Distribution by Venue in {sel_ctr}")
                fig2.update_yaxes(title_text="Economy (RPO)")
                pch(fig2)
                

            # 3. Wickets vs Economy bubble by Venue
            with r1c:
                
                fig3 = go.Figure(go.Scatter(
                    x=bgg2["Economy"], y=bgg2["Wickets"],
                    mode="markers+text",
                    text=bgg2["Venue"],
                    textposition="top center",
                    textfont=dict(family="Poppins",size=9,color=NAVY),
                    marker=dict(
                        size=[max(w*8,14) for w in bgg2["Wickets"]],
                        color=vclrs2,
                        line=dict(color="white",width=2),
                        opacity=0.85),
                ))
                pchart(fig3, 310, f"Bowling Control by Venue in {sel_ctr} (Bubble)")
                fig3.update_xaxes(title_text="Economy (RPO)")
                fig3.update_yaxes(title_text="Wickets")
                pch(fig3)
                

            r2a, r2b = st.columns(2, gap="large")

            # 4. Wickets by Length — Stacked Bar by Venue
            with r2a:
                
                fig4 = go.Figure()
                fig4.add_trace(go.Bar(name="Good Length", x=bgg2["Venue"],
                    y=bgg2["Good"], marker_color=NAVY))
                fig4.add_trace(go.Bar(name="Full Length", x=bgg2["Venue"],
                    y=bgg2["Full"], marker_color=AMBER))
                fig4.add_trace(go.Bar(name="Short", x=bgg2["Venue"],
                    y=bgg2["Short"], marker_color=STEEL))
                fig4.update_layout(barmode="stack")
                pchart(fig4, 310, f"Wickets by Length per Venue in {sel_ctr}")
                fig4.update_xaxes(tickangle=-20)
                fig4.update_yaxes(title_text="Wickets")
                pch(fig4)
                

            # 5. Bowling SR by Venue — Column
            with r2b:
                
                bsr2_clr = [GREEN if b<=20 else (AMBER if b<=30 else RED) for b in bgg2["Bowl_SR"]]
                fig5 = go.Figure(go.Bar(
                    x=bgg2["Venue"], y=bgg2["Bowl_SR"],
                    marker=dict(color=bsr2_clr, line=dict(color="white",width=1)),
                    text=bgg2["Bowl_SR"], textposition="outside",
                    textfont=dict(family="Poppins",size=11),
                ))
                fig5.add_hline(y=25, line_dash="dash", line_color=AMBER,
                    annotation_text="BSR 25",
                    annotation_font=dict(color=AMBER,size=9))
                pchart(fig5, 310, f"Bowling Strike Rate by Venue in {sel_ctr}")
                fig5.update_xaxes(tickangle=-20)
                fig5.update_yaxes(title_text="Balls per Wicket")
                pch(fig5)
                

    # ── ALL-ROUNDER S2 ────────────────────────────────────────
    else:
        b2_ar  = bat_rows(base2)
        bw2_ar = bowl_rows(base2)
        ogg_ar = agg_bat_grp(b2_ar, "Venue") if not b2_ar.empty else pd.DataFrame()
        bgg_ar = agg_bowl_grp(bw2_ar,"Venue") if not bw2_ar.empty else pd.DataFrame()

        r1a, r1b = st.columns(2, gap="large")
        with r1a:
            
            if not ogg_ar.empty and not bgg_ar.empty:
                mg2 = pd.merge(ogg_ar[["Venue","Runs","Average"]],
                               bgg_ar[["Venue","Wickets","Economy"]],
                               on="Venue", how="outer").fillna(0)
                vc_ = [CHART_COLORS[i%len(CHART_COLORS)] for i in range(len(mg2))]
                fig = go.Figure()
                fig.add_trace(go.Bar(name="Runs", x=mg2["Venue"], y=mg2["Runs"],
                    marker=dict(color=vc_, line=dict(color="white",width=1)),
                    text=mg2["Runs"], textposition="outside",
                    textfont=dict(family="Poppins",size=10)))
                fig.add_trace(go.Scatter(name="Wickets", x=mg2["Venue"], y=mg2["Wickets"],
                    mode="lines+markers+text",
                    line=dict(color=AMBER,width=2.5),
                    marker=dict(size=10,color=AMBER,line=dict(color="white",width=2)),
                    text=mg2["Wickets"], textposition="top center",
                    textfont=dict(family="Poppins",size=10,color=AMBER),
                    yaxis="y2"))
                fig.update_layout(
                    yaxis2=dict(overlaying="y",side="right",
                        gridcolor="#f0f3f8",linecolor=BORDER,
                        tickfont=dict(size=9),title_text="Wickets"))
                pchart(fig, 310, f"Bat vs Ball by Venue in {sel_ctr} (Dual Axis)")
                fig.update_xaxes(tickangle=-20)
                fig.update_yaxes(title_text="Runs")
                pch(fig)
            

        with r1b:
            
            if not ogg_ar.empty:
                vc2_ = [CHART_COLORS[i%len(CHART_COLORS)] for i in range(len(ogg_ar))]
                fig2 = go.Figure(go.Bar(
                    y=ogg_ar["Venue"], x=ogg_ar["Runs"], orientation="h",
                    marker=dict(color=vc2_, line=dict(color="white",width=1)),
                    text=ogg_ar["Runs"], textposition="outside"))
                pchart(fig2, 310, f"Runs by Venue in {sel_ctr}")
                fig2.update_xaxes(title_text="Runs")
                pch(fig2)
            

        r2a, r2b = st.columns(2, gap="large")
        with r2a:
            
            if not bgg_ar.empty:
                vc3_ = [CHART_COLORS[i%len(CHART_COLORS)] for i in range(len(bgg_ar))]
                fig3 = go.Figure(go.Bar(
                    x=bgg_ar["Venue"], y=bgg_ar["Wickets"],
                    marker=dict(color=vc3_, line=dict(color="white",width=1)),
                    text=bgg_ar["Wickets"], textposition="outside"))
                pchart(fig3, 310, f"Wickets by Venue in {sel_ctr}")
                fig3.update_xaxes(tickangle=-20)
                fig3.update_yaxes(title_text="Wickets")
                pch(fig3)
            

        with r2b:
            
            if not ogg_ar.empty:
                fig4 = go.Figure(go.Bar(
                    x=ogg_ar["Venue"], y=ogg_ar["Average"],
                    marker=dict(color=[CHART_COLORS[i%len(CHART_COLORS)] for i in range(len(ogg_ar))],
                        line=dict(color="white",width=1)),
                    text=ogg_ar["Average"], textposition="outside"))
                fig4.add_hline(y=50, line_dash="dash", line_color=AMBER,
                    annotation_text="Avg 50",
                    annotation_font=dict(color=AMBER,size=9))
                pchart(fig4, 310, f"Batting Average by Venue in {sel_ctr}")
                fig4.update_xaxes(tickangle=-20)
                fig4.update_yaxes(title_text="Average")
                pch(fig4)
            

# ══════════════════════════════════════════════════════════════
#  ACTION BAR + FOOTER
# ══════════════════════════════════════════════════════════════
yl = f" ({sel_year})" if sel_year != "All" else ""
st.markdown(f"""
<div class="action-bar">
  <div style="display:flex;align-items:center;gap:12px;">
    <a class="act-btn act-report" href="/generate_report" target="_self">
      &nbsp;Generate Report
    </a>
    <span style="color:#8899aa;font-size:.78rem;">
      {sel_player} · {ctr_label}{yl} · {total_m} matches
    </span>
  </div>
  <div>
    <a class="act-btn act-compare" href="/compare" target="_self">
      &nbsp;Compare Players
    </a>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="footer">
  Cricket Player Performance Analysis Dashboard &nbsp;|&nbsp;
  Built with Streamlit
</div>
""", unsafe_allow_html=True)