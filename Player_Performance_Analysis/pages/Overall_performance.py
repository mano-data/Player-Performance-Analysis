import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import openpyxl
import base64, os, math

# ── PAGE CONFIG ──────────────────────────────────────────────
st.set_page_config(
    page_title="Overall Performance | Cricket Analysis",
    page_icon="", layout="wide",
    initial_sidebar_state="collapsed",
)

# ── PATHS ────────────────────────────────────────────────────
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

# ── COLOR PALETTE ────────────────────────────────────────────
NAVY  = "#1B3A6B"
AMBER = "#E8A020"
STEEL = "#4A6FA5"
GREEN = "#1E7A34"
RED   = "#C0392B"
LT_BG = "#f7f9fc"
BORDER= "#dde3ee"

# ── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700;800&display=swap');

html, body, [class*="css"] { font-family:'Inter',sans-serif; background:#fff !important; }
#MainMenu, footer, header { visibility:hidden; }
section[data-testid="stSidebar"] { display:none !important; }

.block-container,
[data-testid="stAppViewBlockContainer"],
[data-testid="block-container"] {
    padding:0 60px 60px 60px !important;
    max-width:1560px !important;
    margin:0 auto !important;
    background:#fff !important;
}
[data-testid="stAppViewBlockContainer"] > div:first-child { padding-top:0 !important; }

/* NAVBAR */
.top-header {
    background:#fff; border-bottom:2px solid #1B3A6B;
    padding:0 60px; height:64px;
    display:flex; align-items:center; justify-content:space-between;
    position:fixed;width:100%; top:0; z-index:999;
    box-shadow:0 2px 12px rgba(27,58,107,.08);
    margin:0 -60px 0 -60px;
}
.header-brand { font-family:'Poppins',sans-serif; font-size:1.2rem; font-weight:700; color:#1B3A6B; }
.header-brand span { color:#E8A020; }
.header-nav { display:flex; gap:8px; align-items:center; }
.h-btn {
    background:#1B3A6B; color:#fff !important; border:none;
    padding:7px 18px; border-radius:6px; font-size:.82rem; font-weight:600;
    text-decoration:none !important; display:inline-block; transition:background .2s;
}
.h-btn:hover { background:#243f7a; }
.h-btn-ghost {
    background:transparent; color:#1B3A6B !important; border:1.5px solid #1B3A6B;
    padding:6px 16px; border-radius:6px; font-size:.82rem; font-weight:600;
    text-decoration:none !important; display:inline-block; transition:background .2s;
}
.h-btn-ghost:hover { background:#f0f4fb; }

/* PAGE TITLE */
.page-title-bar {
    background:#f7f9fc; border-left:4px solid #1B3A6B;
    padding:14px 60px; margin:40px -60px 28px -60px;
}
.page-title-bar h2 { font-family:'Poppins',sans-serif; font-size:1.15rem; font-weight:700; color:#1B3A6B; margin:0; }
.page-title-bar p  { color:#6B7A8D; font-size:.8rem; margin:3px 0 0 0; }

/* FILTER */
.filter-wrap {
    background:#f7f9fc; border:1px solid #dde3ee;
    border-radius:10px; padding:22px 28px; margin-bottom:28px;
}
.filter-label { font-size:.75rem; font-weight:600; color:#1B3A6B; letter-spacing:.8px; text-transform:uppercase; margin-bottom:6px; }

/* Streamlit widget overrides */
.stSelectbox > div > div {
    background:#fff !important; border:1.5px solid #c8d0e0 !important;
    border-radius:7px !important; color:#1B3A6B !important; font-size:.88rem !important;
}
.stSelectbox > div > div:focus-within { border-color:#1B3A6B !important; box-shadow:0 0 0 3px rgba(27,58,107,.1) !important; }
div[data-testid="stButton"] > button {
    background:#1B3A6B !important; color:#fff !important;
    font-family:'Poppins',sans-serif !important; font-weight:700 !important;
    font-size:.9rem !important; letter-spacing:1px !important;
    border:none !important; border-radius:7px !important;
    padding:12px 0 !important; width:100% !important;
}
div[data-testid="stButton"] > button:hover { background:#243f7a !important; }

/* SECTION TITLE */
.sec-title {
    font-family:'Poppins',sans-serif; font-size:1rem; font-weight:700; color:#1B3A6B;
    margin:0 0 14px 0; padding-bottom:8px; border-bottom:2px solid #e8edf5;
    display:flex; align-items:center; gap:8px;
}
.sec-title .dot { width:8px; height:8px; background:#1B3A6B; border-radius:50%; display:inline-block; }

/* PROFILE CARD */
.profile-card {
    background:#fff; border:1px solid #dde3ee; border-radius:12px;
    padding:28px 22px; box-shadow:0 2px 12px rgba(27,58,107,.07);
    position:relative; overflow:hidden;
}
.profile-card::before {
    content:''; position:absolute; top:0; left:0; right:0; height:4px;
    background:linear-gradient(90deg,#1B3A6B 0%,#E8A020 100%);
}
.player-name { font-family:'Poppins',sans-serif; font-size:1.2rem; font-weight:700; color:#1B3A6B; text-align:center; margin:12px 0 4px 0; }
.player-role-badge {
    display:block; width:fit-content; margin:0 auto 18px auto;
    padding:3px 14px; border-radius:20px;
    font-size:.72rem; font-weight:700; letter-spacing:.5px; text-transform:uppercase;
}
.badge-batsman   { background:#EBF3FF; color:#1B3A6B; border:1px solid #c0d0eb; }
.badge-bowler    { background:#FFF0E8; color:#C0501A; border:1px solid #f0c0a0; }
.badge-allrounder{ background:#F0F8EE; color:#2A7A2A; border:1px solid #a8d4a0; }
.info-row { display:flex; justify-content:space-between; align-items:center; padding:9px 0; border-bottom:1px solid #f0f3f8; font-size:.83rem; }
.info-row:last-child { border-bottom:none; }
.info-lbl { color:#6B7A8D; font-weight:500; }
.info-val { color:#1B3A6B; font-weight:600; }

/* RECENT TABLE */
.tbl-wrap { background:#fff; border:1px solid #dde3ee; border-radius:12px; overflow:hidden; box-shadow:0 2px 10px rgba(27,58,107,.06); }
.match-tbl { width:100%; border-collapse:collapse; }
.match-tbl thead tr { background:#1B3A6B; }
.match-tbl thead th { padding:10px 13px; font-size:.7rem; font-weight:600; letter-spacing:.7px; text-transform:uppercase; color:#fff; text-align:left; }
.match-tbl tbody tr { border-bottom:1px solid #f0f3f8; transition:background .15s; }
.match-tbl tbody tr:last-child { border-bottom:none; }
.match-tbl tbody tr:hover { background:#f7f9fc; }
.match-tbl tbody td { padding:9px 13px; font-size:.82rem; color:#2c3e50; }
.badge { padding:3px 10px; border-radius:20px; font-size:.7rem; font-weight:700; }
.b-won  { background:#EAF6EE; color:#1E7E34; border:1px solid #b8dfc4; }
.b-lose { background:#FDE8E8; color:#C0392B; border:1px solid #f0b8b8; }
.b-nr   { background:#F4F4F4; color:#666; border:1px solid #ddd; }

/* KPI CARDS */
.kpi-card {
    background:#fff; border:1px solid #dde3ee; border-radius:10px;
    padding:18px 14px 14px 14px; text-align:center;
    box-shadow:0 2px 8px rgba(27,58,107,.05);
    transition:box-shadow .2s,transform .2s;
    position:relative; overflow:hidden;
}
.kpi-card:hover { box-shadow:0 6px 18px rgba(27,58,107,.12); transform:translateY(-2px); }
.kpi-card::after { content:''; position:absolute; bottom:0; left:0; right:0; height:3px; background:#1B3A6B; }
.kpi-card.hl::after   { background:#E8A020; }
.kpi-card.hl-g::after { background:#1E7A34; }
.kpi-val { font-family:'Poppins',sans-serif; font-size:1.65rem; font-weight:700; color:#1B3A6B; line-height:1; margin-bottom:6px; }
.kpi-lbl { font-size:.7rem; color:#8899aa; font-weight:600; letter-spacing:.6px; text-transform:uppercase; }
.kpi-sub { font-size:.68rem; color:#aab0bc; margin-top:3px; }

/* INSIGHT CARDS */
.insight-wrap { display:flex; flex-direction:column; gap:10px; margin-bottom:24px; }
.insight-card {
    background:#f7f9fc; border:1px solid #dde3ee; border-left:4px solid #E8A020;
    border-radius:8px; padding:12px 16px;
    display:flex; align-items:flex-start; gap:12px;
}
.insight-icon { font-size:1.2rem; flex-shrink:0; margin-top:2px; }
.insight-text { font-size:.84rem; color:#2c3e50; line-height:1.5; }
.insight-text strong { color:#1B3A6B; }

/* CHART CARD */
.chart-card {
    background:#fff; border:1px solid #dde3ee; border-radius:12px;
    padding:20px 18px 10px 18px;
    box-shadow:0 2px 10px rgba(27,58,107,.06); margin-bottom:20px;
}

/* DIVIDER */
.divider { border:none; height:1px; background:#e8edf5; margin:28px 0; }

/* NO DATA */
.no-data { text-align:center; padding:50px; color:#8899aa; font-size:.9rem; }
.no-data .icon { font-size:2.5rem; display:block; margin-bottom:10px; }

/* ACTION BAR */
.action-bar {
    background:#f7f9fc; border-top:1px solid #dde3ee;
    padding:26px 60px; margin:32px -60px 0 -60px;
    display:flex; align-items:center; justify-content:space-between;
}
.act-btn {
    display:inline-flex; align-items:center; gap:8px;
    padding:11px 28px; border-radius:7px;
    font-family:'Poppins',sans-serif; font-weight:700; font-size:.88rem;
    letter-spacing:.5px; cursor:pointer; text-decoration:none !important; border:none;
}
.act-report  { background:#fff; color:#1B3A6B !important; border:2px solid #1B3A6B; }
.act-report:hover  { background:#EBF3FF; }
.act-compare { background:#1B3A6B; color:#fff !important; }
.act-compare:hover { background:#243f7a; }

/* FOOTER */
.footer {
    background:#1B3A6B; color:rgba(255,255,255,.7);
    text-align:center; padding:18px 60px;
    font-size:.8rem; margin:0 -60px -60px -60px;
    letter-spacing:.3px;
}
.footer span { color:#E8A020; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ── DATA LOADING ─────────────────────────────────────────────
@st.cache_data
def load_data():
    master = pd.read_csv(MASTER_PATH)
    master.columns = master.columns.str.strip()
    wb = openpyxl.load_workbook(STATS_PATH)
    skip = {"Match_id","Opponent","Venue","Country","Match_result",
            "Dismissed_method","Dismissed_Length","Dismissed_against",
            "Dismissed_by","Dismissed _by","Match_date"}
    stats = {}
    for sheet in wb.sheetnames:
        ws   = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            stats[sheet.strip()] = pd.DataFrame(); continue
        cols = [str(c).strip() if c else f"_c{i}" for i,c in enumerate(rows[0])]
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

# ── HELPERS ──────────────────────────────────────────────────
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
            f'margin:0 auto;font-family:Poppins,sans-serif;font-size:1.8rem;font-weight:700;color:#1B3A6B;">'
            f'{ini}</div>')

def rbadge(r):
    if r=="Won":  return '<span class="badge b-won">Won</span>'
    if r=="Lose": return '<span class="badge b-lose">Lose</span>'
    return '<span class="badge b-nr">N/R</span>'

def calc_avg(runs_s, not_out_s, innings):
    tot = float(runs_s.sum())
    dis = max(innings - int(not_out_s.sum()), 0)
    return round(tot/dis, 2) if dis > 0 else round(tot, 2)

def pchart(fig, h=320, title=""):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Inter", color="#2c3e50", size=11),
        height=h, margin=dict(l=10,r=10,t=42,b=10),
        title=dict(text=title, font=dict(family="Poppins",size=13,color=NAVY), x=0.02, xanchor="left"),
        legend=dict(bgcolor="rgba(255,255,255,.9)", bordercolor=BORDER, borderwidth=1, font=dict(size=10)),
    )
    fig.update_xaxes(gridcolor="#f0f3f8", linecolor=BORDER, tickfont=dict(size=9), zeroline=False)
    fig.update_yaxes(gridcolor="#f0f3f8", linecolor=BORDER, tickfont=dict(size=9), zeroline=False)
    return fig

def kpi_html(val, lbl, sub="", variant=""):
    return (f'<div class="kpi-card {variant}">'
            f'<div class="kpi-val">{val}</div>'
            f'<div class="kpi-lbl">{lbl}</div>'
            + (f'<div class="kpi-sub">{sub}</div>' if sub else "")
            + '</div>')

SENA = {"South Africa","England","New Zealand","Australia"}
ASIA = {"India","Sri Lanka","Bangladesh","Pakistan","UAE"}

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
  <h2>Overall Performance Analysis</h2>
  <p>Comprehensive career statistics, insights and visual trends for any player</p>
</div>
""", unsafe_allow_html=True)

# ── FILTERS ───────────────────────────────────────────────────
fc1, fc2,= st.columns([4, 3])

with fc1:
    st.markdown('<div class="filter-label">Select Player</div>', unsafe_allow_html=True)
    sel_player = st.selectbox("Player", sorted(stats_dict.keys()), label_visibility="collapsed")

with fc2:
    st.markdown('<div class="filter-label">Select Year</div>', unsafe_allow_html=True)
    df_raw = stats_dict.get(sel_player, pd.DataFrame())
    years  = ["All"] + sorted([int(y) for y in df_raw["Match_date"].dropna().dt.year.unique()]) if not df_raw.empty else ["All"]
    sel_year = st.selectbox("Year", years, label_visibility="collapsed")

st.markdown('<div class="filter-label">&nbsp;</div>', unsafe_allow_html=True)
analyze = st.button("ANALYZE NOW", use_container_width=True)


# ── SESSION STATE ─────────────────────────────────────────────
if analyze:
    st.session_state.update(op_ok=True, op_player=sel_player, op_year=sel_year)

if not st.session_state.get("op_ok"):
    st.markdown('<div class="no-data"><span class="icon"></span>Select a player and year above, then click <b>Analyze Now</b></div>', unsafe_allow_html=True)
    st.stop()

sel_player = st.session_state.get("op_player", sel_player)
sel_year   = st.session_state.get("op_year",   sel_year)
df_raw     = stats_dict.get(sel_player, pd.DataFrame())

# ── PLAYER META ───────────────────────────────────────────────
prow = master_df[master_df["PlayerName"].str.strip() == sel_player]
if not prow.empty:
    pi = prow.iloc[0]
    p_country = pi.get("Country","India"); p_role = pi.get("Role","Batsman")
    p_bat     = pi.get("BattingStyle","RHB"); p_bowl = pi.get("BowlingStyle","-")
    p_dob     = pi.get("DOB","N/A");  p_debut = pi.get("DebutYear","N/A")
    p_img     = pi.get("ImageFileName","")
else:
    p_country="India"; p_role="Batsman"; p_bat="RHB"; p_bowl="-"
    p_dob="N/A"; p_debut="N/A"; p_img=""

rl      = p_role.lower()
is_bat  = "batsman"     in rl
is_bowl = "bowler"      in rl
is_all  = "all-rounder" in rl or "all rounder" in rl
bcls    = "badge-bowler" if is_bowl else ("badge-allrounder" if is_all else "badge-batsman")

# ── YEAR FILTER ───────────────────────────────────────────────
if df_raw.empty:
    st.markdown('<div class="no-data"><span class="icon"></span>No data for this player.</div>', unsafe_allow_html=True); st.stop()

wdf = df_raw.copy()
if sel_year != "All":
    wdf = wdf[wdf["Match_date"].dt.year == int(sel_year)]
wdf = wdf.sort_values("Match_date", ascending=False)

# Active batting / bowling rows
def bat_rows(df):
    return df[df["Batted_Flag"]==1] if "Batted_Flag" in df.columns else df

def bowl_rows(df):
    return df[df["Bowl_Flag"]==1] if "Bowl_Flag" in df.columns else df

bat_df  = bat_rows(wdf)
bowl_df = bowl_rows(wdf)

if bat_df.empty and bowl_df.empty:
    st.markdown(f'<div class="no-data"><span class="icon"></span>No data for {sel_player} in {sel_year}.</div>', unsafe_allow_html=True); st.stop()

# ══════════════════════════════════════════════════════════════
#  SECTION 1 : PROFILE + RECENT PERFORMANCE TABLE
# ══════════════════════════════════════════════════════════════
col_p, col_t = st.columns([1, 2.4], gap="large")

with col_p:
    st.markdown('<div class="sec-title"><span class="dot"></span>Player Profile</div>', unsafe_allow_html=True)
    img_src  = get_img_b64(p_img) if p_img else None
    img_html = (f'<img src="{img_src}" style="width:100px;height:100px;border-radius:50%;'
                f'object-fit:cover;border:3px solid #1B3A6B;display:block;margin:0 auto;">'
                if img_src else initials_avatar(sel_player))
    total_m = len(wdf["Match_id"].dropna().unique()) if "Match_id" in wdf.columns else len(wdf)
    st.markdown(f"""
    <div class="profile-card">
        {img_html}
        <div class="player-name">{sel_player}</div>
        <div class="player-role-badge {bcls}">{p_role}</div>
        <div class="info-row"><span class="info-lbl">Batting Style</span><span class="info-val">{p_bat}</span></div>
        <div class="info-row"><span class="info-lbl">Bowling Style</span><span class="info-val">{p_bowl}</span></div>
        <div class="info-row"><span class="info-lbl">Date of Birth</span><span class="info-val">{p_dob}</span></div>
        <div class="info-row"><span class="info-lbl">Debut Year</span><span class="info-val">{p_debut}</span></div>
        <div class="info-row"><span class="info-lbl">Total Matches</span><span class="info-val">{total_m}</span></div>
    </div>
    """, unsafe_allow_html=True)

with col_t:
    st.markdown('<div class="sec-title"><span class="dot"></span>Recent Performance — Last 5 Matches</div>', unsafe_allow_html=True)

    # ── BATSMAN TABLE ─────────────────────────────────────────
    if is_bat and not is_all:
        last5 = bat_df.head(5)
        hdr   = "<tr><th>Date</th><th>Opponent</th><th>Venue</th><th>Runs (Balls)</th><th>Strike Rate</th><th>4s / 6s</th><th>Result</th></tr>"
        body  = ""
        for _, r in last5.iterrows():
            dt   = r["Match_date"].strftime("%d %b %Y") if pd.notnull(r["Match_date"]) else "N/A"
            runs = int(r.get("Runs",0)); balls = int(r.get("Balls_faced",0))
            no   = int(r.get("Not_out",0)); sr = r.get("Strike_Rate",0)
            sr_s = f"{float(sr):.1f}" if sr else "-"
            f4   = int(r.get("4s",0)); f6 = int(r.get("6s",0))
            run_disp = f"{runs}{'*' if no else ''} ({balls})"
            run_sty  = "font-weight:700;color:#1B3A6B;" if runs>=50 else "color:#2c3e50;"
            body += (f"<tr><td><b>{dt}</b></td><td>{r.get('Opponent','')}</td>"
                     f"<td>{r.get('Venue','')}</td>"
                     f"<td style='{run_sty}'>{run_disp}</td>"
                     f"<td>{sr_s}</td><td>{f4} / {f6}</td>"
                     f"<td>{rbadge(str(r.get('Match_result','')))}</td></tr>")

    # ── BOWLER TABLE ──────────────────────────────────────────
    elif is_bowl and not is_all:
        last5 = bowl_df.head(5)
        hdr   = "<tr><th>Date</th><th>Opponent</th><th>Venue</th><th>Overs</th><th>Figures</th><th>Economy</th><th>Result</th></tr>"
        body  = ""
        for _, r in last5.iterrows():
            dt    = r["Match_date"].strftime("%d %b %Y") if pd.notnull(r["Match_date"]) else "N/A"
            balls = int(r.get("Balls_Bowled",0))
            overs = f"{balls//6}.{balls%6}"
            wkts  = int(r.get("Wickets",0)); rc = int(r.get("Runs_Conceded",0))
            eco   = r.get("Economy",0); eco_s = f"{float(eco):.2f}" if eco else "-"
            fig_s = f"{wkts}-{rc}"
            fig_sty = "font-weight:700;color:#1E7A34;" if wkts>=3 else ("font-weight:600;color:#1B3A6B;" if wkts>=1 else "")
            body += (f"<tr><td><b>{dt}</b></td><td>{r.get('Opponent','')}</td>"
                     f"<td>{r.get('Venue','')}</td>"
                     f"<td style='font-weight:600;color:#1B3A6B;'>{overs}</td>"
                     f"<td style='{fig_sty}'>{fig_s}</td>"
                     f"<td>{eco_s}</td>"
                     f"<td>{rbadge(str(r.get('Match_result','')))}</td></tr>")

    # ── ALL-ROUNDER TABLE ─────────────────────────────────────
    else:
        # merge bat + bowl on match date / match_id
        bat5  = bat_df.head(5).copy()
        body  = ""
        hdr   = "<tr><th>Date</th><th>Opponent</th><th>Venue</th><th>Runs* (Balls)</th><th>SR</th><th>Figures</th><th>Economy</th><th>Result</th></tr>"
        for _, r in bat5.iterrows():
            dt   = r["Match_date"].strftime("%d %b %Y") if pd.notnull(r["Match_date"]) else "N/A"
            runs = int(r.get("Runs",0)); balls = int(r.get("Balls_faced",0))
            no   = int(r.get("Not_out",0)); sr = r.get("Strike_Rate",0)
            sr_s = f"{float(sr):.1f}" if sr else "-"
            run_disp = f"{runs}{'*' if no else ''} ({balls})"
            # bowling from same row (AR has both)
            bf = r.get("Bowl_Flag",0)
            if bf:
                bb   = int(r.get("Balls_Bowled",0))
                wkts = int(r.get("Wickets",0)); rc2 = int(r.get("Runs_Conceded",0))
                eco2 = r.get("Economy",0); eco_s = f"{float(eco2):.2f}" if eco2 else "-"
                fig_s= f"{wkts}-{rc2}"
            else:
                fig_s="-"; eco_s="-"
            run_sty = "font-weight:700;color:#1B3A6B;" if runs>=50 else "color:#2c3e50;"
            body += (f"<tr><td><b>{dt}</b></td><td>{r.get('Opponent','')}</td>"
                     f"<td>{r.get('Venue','')}</td>"
                     f"<td style='{run_sty}'>{run_disp}</td>"
                     f"<td>{sr_s}</td><td>{fig_s}</td><td>{eco_s}</td>"
                     f"<td>{rbadge(str(r.get('Match_result','')))}</td></tr>")

    st.markdown(
        f'<div class="tbl-wrap"><table class="match-tbl"><thead>{hdr}</thead><tbody>{body}</tbody></table></div>',
        unsafe_allow_html=True
    )

# ══════════════════════════════════════════════════════════════
#  SECTION 2 : KPI CARDS
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="sec-title"><span class="dot"></span>Key Performance Indicators</div>', unsafe_allow_html=True)

if is_bat and not is_all:
    inn     = len(bat_df)
    tot_r   = int(bat_df["Runs"].sum())
    not_out = bat_df.get("Not_out", pd.Series(0, index=bat_df.index))
    avg     = calc_avg(bat_df["Runs"], not_out, inn)
    avg_sr  = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(), 1)
    tot_4s  = int(bat_df["4s"].sum()); tot_6s = int(bat_df["6s"].sum())
    fif     = int(bat_df["50s"].sum()) if "50s" in bat_df else 0
    huns    = int(bat_df["100s"].sum()) if "100s" in bat_df else 0
    hs_r    = bat_df.loc[bat_df["Runs"].idxmax()]
    hs      = f"{int(hs_r['Runs'])}{'*' if int(hs_r.get('Not_out',0)) else ''}"
    dot_pct = round(bat_df["Dots"].sum()/max(bat_df["Balls_faced"].sum(),1)*100,1)
    kpis = [
        (str(inn),            "Innings",     "Matches batted",   ""),
        (str(tot_r),          "Total Runs",  "Career total",     "hl"),
        (str(avg),            "Average",     "Runs/dismissal",   ""),
        (str(avg_sr),         "Strike Rate", "Avg SR",           ""),
        (f"{tot_4s}/{tot_6s}","4s / 6s",     "Boundaries",       "hl"),
        (f"{fif}/{huns}",     "50s / 100s",  "Milestones",       "hl-g"),
        (hs,                  "Best Score",  "Highest innings",  ""),
    ]

elif is_bowl and not is_all:
    inn     = len(bowl_df)
    t_balls = int(bowl_df["Balls_Bowled"].sum())
    overs_s = f"{t_balls//6}.{t_balls%6}"
    t_wkts  = int(bowl_df["Wickets"].sum())
    avg_eco = round(bowl_df["Economy"].replace(0,float("nan")).mean(), 2)
    bsr     = round(t_balls/max(t_wkts,1), 1)
    t3f     = int((bowl_df["Wickets"]>=3).sum())
    t5f     = int((bowl_df["Wickets"]>=5).sum())
    dots_b  = int(bowl_df["Dot_Balls"].sum()) if "Dot_Balls" in bowl_df else 0
    if t_wkts > 0:
        bf   = bowl_df.sort_values(["Wickets","Runs_Conceded"],ascending=[False,True]).iloc[0]
        bfig = f"{int(bf['Wickets'])}/{int(bf.get('Runs_Conceded',0))}"
    else: bfig="0/0"
    kpis = [
        (str(inn),       "Matches",      "Career",          ""),
        (overs_s,        "Overs",        "Total bowled",    ""),
        (str(t_wkts),    "Wickets",      "Career total",    "hl"),
        (str(avg_eco),   "Economy",      "Runs/over",       ""),
        (str(bsr),       "Bowling SR",   "Balls/wicket",    ""),
        (f"{t3f}/{t5f}", "3-Fer/5-Fer",  "Wicket hauls",    "hl-g"),
        (str(dots_b),    "Dot Balls",    "Scoreless balls", ""),
        (bfig,           "Best Figures", "Wkts/Runs",       "hl"),
    ]

else:  # All-rounder
    b_inn   = len(bat_df); bo_inn = len(bowl_df)
    tot_r   = int(bat_df["Runs"].sum()) if not bat_df.empty else 0
    not_out = bat_df.get("Not_out",pd.Series(0,index=bat_df.index)) if not bat_df.empty else pd.Series([0])
    avg     = calc_avg(bat_df["Runs"] if not bat_df.empty else pd.Series([0]), not_out, max(b_inn,1))
    avg_sr  = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(),1) if not bat_df.empty else 0
    fif     = int(bat_df["50s"].sum()) if not bat_df.empty and "50s" in bat_df else 0
    huns    = int(bat_df["100s"].sum()) if not bat_df.empty and "100s" in bat_df else 0
    t_wkts  = int(bowl_df["Wickets"].sum()) if not bowl_df.empty else 0
    avg_eco = round(bowl_df["Economy"].replace(0,float("nan")).mean(),2) if not bowl_df.empty else 0
    bsr_v   = round(int(bowl_df["Balls_Bowled"].sum())/max(t_wkts,1),1) if not bowl_df.empty else 0
    kpis = [
        (str(b_inn),         "Innings",     "Matches batted",   ""),
        (str(tot_r),         "Total Runs",  "Career",           "hl"),
        (str(avg),           "Bat Average", "Runs/dismissal",   ""),
        (str(avg_sr),        "Strike Rate", "Avg SR",           ""),
        (f"{fif}/{huns}",    "50s / 100s",  "Milestones",       "hl-g"),
        (str(t_wkts),        "Wickets",     "Career",           "hl"),
        (str(avg_eco),       "Economy",     "Runs/over",        ""),
        (str(bsr_v),         "Bowl SR",     "Balls/wicket",     ""),
    ]

# Render KPIs
cols = st.columns(len(kpis))
for i,(v,l,s,var) in enumerate(kpis):
    with cols[i]:
        st.markdown(kpi_html(v,l,s,var), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SECTION 3 : KEY INSIGHTS
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="sec-title"><span class="dot"></span>Key Insights</div>', unsafe_allow_html=True)

insights = []

if is_bat and not is_all:
    # Insight 1: Win % when score ≥ 50
    hi_bat = bat_df[bat_df["Runs"]>=50]
    if not hi_bat.empty:
        wp = round((hi_bat["Match_result"]=="Won").sum()/len(hi_bat)*100,1)
        insights.append(( f"Team wins <strong>{wp}%</strong> of matches when {sel_player} scores 50+ "
                         f"({int((hi_bat['Match_result']=='Won').sum())} wins from {len(hi_bat)} such innings)"))

    # Insight 2: Boundary contribution %
    bnd_runs = int(bat_df["4s"].sum()*4 + bat_df["6s"].sum()*6)
    tot_r2   = int(bat_df["Runs"].sum())
    if tot_r2 > 0:
        bpct = round(bnd_runs/tot_r2*100,1)
        insights.append(( f"<strong>{bpct}%</strong> of runs come from boundaries "
                         f"({bnd_runs} boundary runs out of {tot_r2} total)"))

    # Insight 3: Asia vs SENA
    asia_df = bat_df[bat_df["Country"].isin(ASIA)] if "Country" in bat_df else pd.DataFrame()
    sena_df = bat_df[bat_df["Country"].isin(SENA)] if "Country" in bat_df else pd.DataFrame()
    if not asia_df.empty and not sena_df.empty:
        a_avg = calc_avg(asia_df["Runs"], asia_df.get("Not_out",pd.Series(0,index=asia_df.index)), len(asia_df))
        s_avg = calc_avg(sena_df["Runs"], sena_df.get("Not_out",pd.Series(0,index=sena_df.index)), len(sena_df))
        loc   = "Asia" if a_avg >= s_avg else "SENA"
        insights.append(( f"{sel_player} performs better in <strong>{loc}</strong> — "
                         f"Asia Avg: <strong>{a_avg}</strong> | SENA Avg: <strong>{s_avg}</strong>"))

    # Insight 4: 50→100 conversion
    if fif + huns > 0:
        conv = round(huns/(fif+huns)*100,1)
        insights.append(( f"50 to 100 conversion rate: <strong>{conv}%</strong> "
                         f"({huns} centuries from {fif+huns} fifty-plus scores)"))

elif is_bowl and not is_all:
    # Insight 1: Win % when taking 3+ wickets
    haul3 = bowl_df[bowl_df["Wickets"]>=3]
    if not haul3.empty:
        wp = round((haul3["Match_result"]=="Won").sum()/len(haul3)*100,1)
        insights.append(( f"Team wins <strong>{wp}%</strong> of matches when {sel_player} takes 3+ wickets"))

    # Insight 2: Dot ball efficiency
    if t_balls > 0:
        dot_eff = round(dots_b/t_balls*100,1)
        insights.append(( f"<strong>{dot_eff}%</strong> of balls bowled are dot balls ({dots_b} out of {t_balls} balls)"))

    # Insight 3: RHB vs LHB
    if "Wicket_RHB" in bowl_df and "Wicket_LHB" in bowl_df:
        rhb_w = int(bowl_df["Wicket_RHB"].sum()); lhb_w = int(bowl_df["Wicket_LHB"].sum())
        better = "Right-Handed" if rhb_w >= lhb_w else "Left-Handed"
        insights.append(( f"More effective vs <strong>{better} batters</strong> — RHB: {rhb_w} wkts | LHB: {lhb_w} wkts"))

    # Insight 4: Economy in wins vs losses
    wins_eco  = bowl_df[bowl_df["Match_result"]=="Won"]["Economy"].replace(0,float("nan")).mean()
    loss_eco  = bowl_df[bowl_df["Match_result"]=="Lose"]["Economy"].replace(0,float("nan")).mean()
    if pd.notnull(wins_eco) and pd.notnull(loss_eco):
        insights.append(( f"Economy in <strong>wins: {wins_eco:.2f}</strong> vs <strong>losses: {loss_eco:.2f}</strong> — "
                         f"{'tighter in wins ' if wins_eco < loss_eco else 'more expensive in wins'}"))

else:  # All-rounder
    if not bat_df.empty:
        wp5 = round((bat_df[bat_df["Runs"]>=30]["Match_result"]=="Won").sum()/max(len(bat_df[bat_df["Runs"]>=30]),1)*100,1)
        insights.append(( f"Team wins <strong>{wp5}%</strong> of matches when {sel_player} scores 30+"))
    if not bowl_df.empty:
        haul2 = bowl_df[bowl_df["Wickets"]>=2]
        if not haul2.empty:
            wp2 = round((haul2["Match_result"]=="Won").sum()/len(haul2)*100,1)
            insights.append(( f"Team wins <strong>{wp2}%</strong> of matches when {sel_player} takes 2+ wickets"))
    if not bat_df.empty:
        bnd_r = int(bat_df["4s"].sum()*4 + bat_df["6s"].sum()*6)
        tr2   = int(bat_df["Runs"].sum())
        if tr2 > 0:
            insights.append(( f"<strong>{round(bnd_r/tr2*100,1)}%</strong> of batting runs come from boundaries"))
    if not bowl_df.empty and "Wicket_RHB" in bowl_df.columns:
        rw = int(bowl_df["Wicket_RHB"].sum()); lw = int(bowl_df["Wicket_LHB"].sum())
        insights.append(( f"Bowling: <strong>{rw}</strong> wickets vs RHB | <strong>{lw}</strong> wickets vs LHB"))

if insights:
    st.markdown('<div class="insight-wrap">', unsafe_allow_html=True)
    for txt in insights:
        st.markdown(
            f'<div class="insight-card"><span class="insight-icon"></span>'
            f'<span class="insight-text">{txt}</span></div>',
            unsafe_allow_html=True
        )
    st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="no-data">No insights available for selected filters.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SECTION 4 : CHARTS
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
chart_mode = "Year-wise" if sel_year == "All" else f"{sel_year} Month-wise"
st.markdown(f'<div class="sec-title"><span class="dot"></span>Analysis Charts — {chart_mode}</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
#  WAGON WHEEL HELPER — CricViz-style, 8 equal sectors
#  Ported from Strength_weakness.py for consistency
# ─────────────────────────────────────────────────────────────
def make_wagon_wheel(df, title, player_name="", h=420):
    """
    CricViz/ESPNcricinfo-style 8-sector wagon wheel.
    • 8 × 45° sectors, OFF-SIDE left, LEG-SIDE right
    • Colour ramp: #D6E8F7 (low) → #1B3A6B (high)
    • Large bold run values at 62% radius
    • Zone labels at 84% radius (white on dark, steel on light)
    • Colour-scale legend strip below wheel
    """
    ZONES = [
        #  zone_label       data_col            centre_angle_deg
        ("Long Off",     "Runs_Longoff",         112.5),
        ("Cover",        "Runs_Covers",          157.5),
        ("Point",        "Runs_Point",           202.5),
        ("Third Man",    "Runs_Thirdman",        247.5),
        ("Fine Leg",     "Runs_Fineleg",         292.5),
        ("Square Leg",   "Runs_Square",          337.5),
        ("Mid Wicket",   "Runs_Midwct",           22.5),
        ("Long On",      "Runs_Longon",           67.5),
    ]

    N          = len(ZONES)
    SECTOR_DEG = 360.0 / N     # 45° each
    ARC_PTS    = 80             # arc smoothness
    R          = 1.0            # wheel radius
    CX, CY     = 0.0, 0.0      # centre

    # ── aggregate per-zone run totals ─────────────────────────
    vals = [int(df[col].sum()) if col in df.columns else 0
            for _, col, _ in ZONES]

    total_runs = sum(vals)
    max_v      = max(vals) if max(vals) > 0 else 1

    # ── colour: #D6E8F7 (low) → #1B3A6B (high) ───────────────
    def fill_color(v):
        t = v / max_v
        r = int(214 + ( 27 - 214) * t)
        g = int(232 + ( 58 - 232) * t)
        b = int(247 + (107 - 247) * t)
        return f"rgb({r},{g},{b})"

    def num_color(v):
        return "#ffffff" if v / max_v >= 0.55 else "#1B3A6B"

    def rim_color(v):
        return "rgba(255,255,255,0.90)" if v / max_v >= 0.55 else "#4A6FA5"

    # ── build figure ──────────────────────────────────────────
    fig = go.Figure()

    for i, (zone_lbl, col, ctr_ang) in enumerate(ZONES):
        v  = vals[i]
        a0 = ctr_ang - SECTOR_DEG / 2.0
        a1 = ctr_ang + SECTOR_DEG / 2.0

        # Build polygon: origin → arc points → back to origin
        arc_xs = [CX + R * math.cos(math.radians(a0 + (a1 - a0) * k / ARC_PTS))
                  for k in range(ARC_PTS + 1)]
        arc_ys = [CY + R * math.sin(math.radians(a0 + (a1 - a0) * k / ARC_PTS))
                  for k in range(ARC_PTS + 1)]
        poly_x = [CX] + arc_xs + [CX]
        poly_y = [CY] + arc_ys + [CY]

        # ① filled sector
        fig.add_trace(go.Scatter(
            x=poly_x, y=poly_y,
            fill="toself",
            fillcolor=fill_color(v),
            line=dict(color="white", width=2.8),
            mode="lines",
            hovertemplate=f"<b>{zone_lbl}</b><br>{v} runs<extra></extra>",
            showlegend=False,
        ))

        # ② large run-value number at 62% radius
        num_r = R * 0.62
        nx = CX + num_r * math.cos(math.radians(ctr_ang))
        ny = CY + num_r * math.sin(math.radians(ctr_ang))
        fig.add_annotation(
            x=nx, y=ny,
            text=f"<b>{v}</b>",
            showarrow=False,
            font=dict(family="Poppins", size=20, color=num_color(v)),
            xanchor="center", yanchor="middle",
        )

        # ③ zone label at 84% radius
        rim_r = R * 0.84
        rx = CX + rim_r * math.cos(math.radians(ctr_ang))
        ry = CY + rim_r * math.sin(math.radians(ctr_ang))
        fig.add_annotation(
            x=rx, y=ry,
            text=f"<span style='font-size:8px'>{zone_lbl}</span>",
            showarrow=False,
            font=dict(family="Inter", size=8, color=rim_color(v)),
            xanchor="center", yanchor="middle",
        )

    # ── OFF-SIDE / LEG-SIDE midline labels ────────────────────
    for lx_, txt_ in [(-0.46, "OFF-SIDE"), (0.46, "LEG-SIDE")]:
        fig.add_annotation(
            x=lx_, y=0.0,
            text=f"<b>{txt_}</b>",
            showarrow=False,
            font=dict(family="Inter", size=9, color="rgba(255,255,255,0.75)"),
            xanchor="center", yanchor="middle",
        )

    # ── colour-scale legend strip ─────────────────────────────
    STRIP_Y0, STRIP_Y1 = -1.18, -1.08
    for s in range(30):
        t  = s / 29
        r_ = int(214 + ( 27 - 214) * t)
        g_ = int(232 + ( 58 - 232) * t)
        b_ = int(247 + (107 - 247) * t)
        x0 = -0.85 + s * (1.70 / 30)
        x1 = x0 + 1.70 / 30
        fig.add_shape(
            type="rect",
            x0=x0, x1=x1, y0=STRIP_Y0, y1=STRIP_Y1,
            fillcolor=f"rgb({r_},{g_},{b_})",
            line_width=0,
        )
    fig.add_annotation(x=-0.86, y=-1.24, text="Fewer runs",
        showarrow=False,
        font=dict(family="Inter", size=8, color="#8899aa"),
        xanchor="left")
    fig.add_annotation(x=0.86, y=-1.24, text="Most runs",
        showarrow=False,
        font=dict(family="Inter", size=8, color=NAVY),
        xanchor="right")

    # ── player name + total subtitle ─────────────────────────
    if player_name:
        fig.add_annotation(
            x=0, y=1.20,
            text=(f"<span style='font-size:10.5px;color:#6B7A8D'>"
                  f"{player_name}&nbsp;&nbsp;•&nbsp;&nbsp;"
                  f"Total runs: <b>{total_runs}</b></span>"),
            showarrow=False,
            font=dict(family="Inter", size=11, color="#6B7A8D"),
            xanchor="center",
        )

    # ── layout ────────────────────────────────────────────────
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=h,
        margin=dict(l=10, r=10, t=44, b=48),
        title=dict(
            text=title,
            font=dict(family="Poppins", size=13, color=NAVY),
            x=0.02, xanchor="left",
        ),
        xaxis=dict(visible=False, range=[-1.30, 1.30]),
        yaxis=dict(
            visible=False,
            range=[-1.38, 1.30],
            scaleanchor="x",
            scaleratio=1,
        ),
    )
    return fig

# ─────────────────────────────────────────────────────────────
#  BATSMAN CHARTS
# ─────────────────────────────────────────────────────────────
if is_bat and not is_all:
    # Prepare time axis
    year_all = sel_year == "All"
    if year_all:
        bat_df["_grp"] = bat_df["Match_date"].dt.year.astype(str)
        grp_label = "Year"
    else:
        bat_df["_grp"] = bat_df["Match_date"].dt.strftime("%b")
        month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        bat_df["_grp"] = pd.Categorical(bat_df["_grp"], categories=month_order, ordered=True)
        grp_label = "Month"

    grp = bat_df.groupby("_grp", observed=True).agg(
        Runs=("Runs","sum"), Innings=("Runs","count"),
        SR=("Strike_Rate","mean"),
        Fours=("4s","sum"), Sixes=("6s","sum"),
        Dots=("Dots","sum"), Singles=("1s_2s_3s","sum"),
        Not_out=("Not_out","sum"),
    ).reset_index()
    grp["Average"] = grp.apply(lambda r: round(r["Runs"]/max(r["Innings"]-r["Not_out"],1),2), axis=1)
    grp["SR"]      = grp["SR"].round(1)
    x = grp["_grp"].astype(str)

    # ROW 1 — 3 charts
    r1c1, r1c2, r1c3 = st.columns(3, gap="large")

    with r1c1:
        fig1 = go.Figure(go.Bar(
            x=x, y=grp["Runs"],
            marker=dict(color=NAVY, line=dict(color="white",width=1)),
            text=grp["Runs"], textposition="outside",
            textfont=dict(family="Poppins",size=10,color=NAVY),
        ))
        pchart(fig1, 300, f"Runs Distribution ({grp_label}-wise)")
        fig1.update_xaxes(tickangle=-25)
        fig1.update_yaxes(title_text="Runs")
        fig1.update_yaxes(
        range=[0, grp["Runs"].max() * 1.20]
        )
        st.plotly_chart(fig1, use_container_width=True)

    with r1c2:
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=grp["Average"], y=grp["SR"],
            mode="markers+text",
            text=x,
            textposition="top center",
            textfont=dict(family="Poppins",size=9),
            marker=dict(size=12, color=NAVY,
                line=dict(color=AMBER,width=2)),
            name="Avg vs SR",
        ))
        pchart(fig2, 300, f"Strike Rate vs Average ({grp_label}-wise)")
        fig2.update_xaxes(title_text="Average")
        fig2.update_yaxes(title_text="Strike Rate")
        st.plotly_chart(fig2, use_container_width=True)

    with r1c3:
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(name="4s Runs", x=x, y=grp["Fours"]*4, marker_color=NAVY))
        fig3.add_trace(go.Bar(name="6s Runs", x=x, y=grp["Sixes"]*6, marker_color=AMBER))
        fig3.add_trace(go.Bar(name="Dot Balls",x=x, y=grp["Dots"],   marker_color="#E8E8E8"))
        fig3.add_trace(go.Bar(name="1s/2s/3s", x=x, y=grp["Singles"],marker_color=STEEL))
        fig3.update_layout(barmode="stack")
        pchart(fig3, 300, "Boundary Contributions")
        fig3.update_xaxes(tickangle=-25)
        st.plotly_chart(fig3, use_container_width=True,)

    # ROW 2 — Wagon Wheel + Dismissal
    r2c1, r2c2 = st.columns(2, gap="large")

    with r2c1:
        fig_ww = make_wagon_wheel(bat_df, "Scoring Area — Wagon Wheel")
        st.plotly_chart(fig_ww, use_container_width=True)

    with r2c2:
        dis = bat_df[bat_df["Dismissed_method"].astype(str).isin(
            [m for m in bat_df["Dismissed_method"].unique() if str(m) not in ["Not Applicable","0","nan",""]]
        )]
        if not dis.empty:
            dm = dis.groupby("Dismissed_method").size().reset_index(name="Count")
            colors_pie = [NAVY, AMBER, STEEL, GREEN, RED, "#7B5EA7", "#E8A020", "#2ECC71"]
            fig4 = go.Figure(go.Pie(
                labels=dm["Dismissed_method"], values=dm["Count"],
                hole=0.45,
                marker=dict(colors=colors_pie[:len(dm)], line=dict(color="white",width=2)),
                textfont=dict(family="Poppins",size=12),
                textinfo="label+percent",
            ))
            fig4.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", height=370,
                margin=dict(l=10,r=10,t=42,b=10),
                title=dict(text="Dismissal Analysis", font=dict(family="Poppins",size=13,color=NAVY), x=0.02),
                showlegend=True,
                legend=dict(font=dict(size=10), bgcolor="rgba(255,255,255,.9)"),
            )
            st.plotly_chart(fig4, use_container_width=True)
        else:
            st.markdown('<div class="no-data">No dismissal data.</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
#  BOWLER CHARTS
# ─────────────────────────────────────────────────────────────
elif is_bowl and not is_all:
    year_all = sel_year == "All"
    if year_all:
        bowl_df["_grp"] = bowl_df["Match_date"].dt.year.astype(str)
        grp_label = "Year"
    else:
        bowl_df["_grp"] = bowl_df["Match_date"].dt.strftime("%b")
        month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        bowl_df["_grp"] = pd.Categorical(bowl_df["_grp"], categories=month_order, ordered=True)
        grp_label = "Month"

    bgrp = bowl_df.groupby("_grp", observed=True).agg(
        Wickets=("Wickets","sum"),
        Economy=("Economy","mean"),
        Balls=("Balls_Bowled","sum"),
        Dots=("Dot_Balls","sum"),
    ).reset_index()
    bgrp["Bowl_SR"] = (bgrp["Balls"] / bgrp["Wickets"].replace(0,float("nan"))).round(1)
    bgrp["Economy"] = bgrp["Economy"].round(2)
    x = bgrp["_grp"].astype(str)

    # 2x2 layout
    r1c1, r1c2 = st.columns(2, gap="large")
    with r1c1:
        fig1 = go.Figure(go.Bar(
            x=x, y=bgrp["Wickets"],
            marker=dict(color=[GREEN if w>=3 else NAVY for w in bgrp["Wickets"]],
                line=dict(color="white",width=1)),
            text=bgrp["Wickets"], textposition="outside",
            textfont=dict(family="Poppins",size=11,color=NAVY),
        ))
        pchart(fig1, 310, f"Wickets ({grp_label}-wise)")
        fig1.update_xaxes(tickangle=-25)
        st.plotly_chart(fig1, use_container_width=True)

    with r1c2:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=bgrp["Economy"], y=bgrp["Bowl_SR"].fillna(0),
            mode="markers+text", text=x,
            textposition="top center", textfont=dict(family="Poppins",size=9),
            marker=dict(size=12, color=NAVY, line=dict(color=AMBER,width=2)),
        ))
        pchart(fig2, 310, f"Economy vs Bowling SR ({grp_label}-wise)")
        fig2.update_xaxes(title_text="Economy")
        fig2.update_yaxes(title_text="Bowling Strike Rate")
        st.plotly_chart(fig2, use_container_width=True)

    r2c1, r2c2 = st.columns(2, gap="large")
    with r2c1:
        # Wickets by Length
        full_d  = int(bowl_df["Full_Dismissals"].sum()) if "Full_Dismissals" in bowl_df else 0
        good_d  = int(bowl_df.get("Good_Dissmisals", bowl_df.get("Good_Dismissals", pd.Series([0]))).sum())
        short_d = int(bowl_df.get("Short_Dissmisals",bowl_df.get("Short_Dismissals",pd.Series([0]))).sum())
        fig3 = go.Figure(go.Bar(
            x=["Good Length","Full Length","Short Length"],
            y=[good_d, full_d, short_d],
            marker=dict(color=[NAVY, AMBER, STEEL], line=dict(color="white",width=1)),
            text=[good_d, full_d, short_d], textposition="outside",
            textfont=dict(family="Poppins",size=12,color=NAVY),
        ))
        pchart(fig3, 310, "Wickets by Ball Length")
        fig3.update_yaxes(title_text="Wickets")
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        rhb_r  = int(bowl_df["Runs_RHB"].sum()) if "Runs_RHB" in bowl_df else 0
        lhb_r  = int(bowl_df["Runs_LHB"].sum()) if "Runs_LHB" in bowl_df else 0
        rhb_w  = int(bowl_df["Wicket_RHB"].sum()) if "Wicket_RHB" in bowl_df else 0
        lhb_w  = int(bowl_df["Wicket_LHB"].sum()) if "Wicket_LHB" in bowl_df else 0
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(name="Runs vs RHB",   x=["RHB","LHB"], y=[rhb_r, lhb_r], marker_color=NAVY))
        fig4.add_trace(go.Bar(name="Wickets vs RHB", x=["RHB","LHB"], y=[rhb_w, lhb_w], marker_color=AMBER))
        fig4.update_layout(barmode="group")
        pchart(fig4, 310, "RHB vs LHB Performance")
        st.plotly_chart(fig4, use_container_width=True)

# ─────────────────────────────────────────────────────────────
#  ALL-ROUNDER CHARTS
# ─────────────────────────────────────────────────────────────
else:
    year_all = sel_year == "All"
    if year_all:
        bat_df["_grp"]  = bat_df["Match_date"].dt.year.astype(str)
        bowl_df["_grp"] = bowl_df["Match_date"].dt.year.astype(str)
        grp_label = "Year"
    else:
        month_order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        bat_df["_grp"]  = pd.Categorical(bat_df["Match_date"].dt.strftime("%b"),  categories=month_order, ordered=True)
        bowl_df["_grp"] = pd.Categorical(bowl_df["Match_date"].dt.strftime("%b"), categories=month_order, ordered=True)
        grp_label = "Month"

    r1c1, r1c2 = st.columns(2, gap="large")
    with r1c1:
        # Bat vs Ball Impact scatter (per match)
        merged = bat_df[["Match_date","Runs"]].copy()
        merged = merged.join(bowl_df[["Match_date","Wickets"]].set_index("Match_date"), on="Match_date", how="left")
        merged["Wickets"] = merged["Wickets"].fillna(0).astype(int)
        fig1 = go.Figure(go.Scatter(
            x=merged["Runs"], y=merged["Wickets"],
            mode="markers",
            marker=dict(size=10, color=NAVY, opacity=0.75, line=dict(color=AMBER,width=1.5)),
            text=merged["Match_date"].dt.strftime("%d %b %Y"),
            hovertemplate="<b>%{text}</b><br>Runs: %{x}<br>Wickets: %{y}<extra></extra>",
        ))
        pchart(fig1, 310, "Bat vs Ball Impact (Runs vs Wickets per Match)")
        fig1.update_xaxes(title_text="Runs Scored")
        fig1.update_yaxes(title_text="Wickets Taken")
        st.plotly_chart(fig1, use_container_width=True)

    with r1c2:
        # Runs Distribution year/month
        bgrp2 = bat_df.groupby("_grp", observed=True).agg(Runs=("Runs","sum")).reset_index()
        fig2 = go.Figure(go.Bar(
            x=bgrp2["_grp"].astype(str), y=bgrp2["Runs"],
            marker=dict(color=NAVY, line=dict(color="white",width=1)),
            text=bgrp2["Runs"], textposition="outside",
            textfont=dict(family="Poppins",size=10,color=NAVY),
        ))
        pchart(fig2, 310, f"Runs Distribution ({grp_label}-wise)")
        fig2.update_xaxes(tickangle=-25); fig2.update_yaxes(title_text="Runs")
        st.plotly_chart(fig2, use_container_width=True)

    r2c1, r2c2 = st.columns(2, gap="large")
    with r2c1:
        # Wickets distribution
        wgrp = bowl_df.groupby("_grp", observed=True).agg(Wickets=("Wickets","sum")).reset_index()
        fig3 = go.Figure(go.Bar(
            x=wgrp["_grp"].astype(str), y=wgrp["Wickets"],
            marker=dict(color=[GREEN if w>=3 else AMBER for w in wgrp["Wickets"]],
                line=dict(color="white",width=1)),
            text=wgrp["Wickets"], textposition="outside",
            textfont=dict(family="Poppins",size=11),
        ))
        pchart(fig3, 310, f"Wickets ({grp_label}-wise)")
        fig3.update_xaxes(tickangle=-25); fig3.update_yaxes(title_text="Wickets")
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        fig4 = make_wagon_wheel(bat_df, "Scoring Area — Wagon Wheel")
        st.plotly_chart(fig4, use_container_width=True)

# ══════════════════════════════════════════════════════════════
#  BOTTOM ACTION BAR
# ══════════════════════════════════════════════════════════════
yl = f" ({sel_year})" if sel_year != "All" else ""
st.markdown(f"""
<div class="action-bar">
  <div style="display:flex;align-items:center;gap:12px;">
    <a class="act-btn act-report" href="/generate_report" target="_self">&nbsp;Generate Report</a>
    <span style="color:#8899aa;font-size:.78rem;">{sel_player}{yl} · {total_m} matches</span>
  </div>
  <div>
    <a class="act-btn act-compare" href="/compare" target="_self">&nbsp;Compare Players</a>
  </div>
</div>
""", unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  Cricket Player Performance Analysis &nbsp;|&nbsp; Built with Streamlit
</div>
""", unsafe_allow_html=True)