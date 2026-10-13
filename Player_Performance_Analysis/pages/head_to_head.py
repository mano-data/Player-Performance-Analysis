import streamlit as st
import pandas as pd
import openpyxl, base64, os, io, re
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, HRFlowable, KeepTogether, PageBreak
)

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Comparison Report | Cricket Analytics",
    page_icon="", layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
#  PATHS
# ─────────────────────────────────────────────
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

# ─────────────────────────────────────────────
#  BRAND PALETTE
# ─────────────────────────────────────────────
NAVY  = "#1B3A6B"
AMBER = "#E8A020"
STEEL = "#4A6FA5"
GREEN = "#1E7A34"
RED   = "#C0392B"
LIGHT = "#EBF3FF"
BGRAY = "#f7f9fc"

# ─────────────────────────────────────────────
#  GLOBAL CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700;800&display=swap');

html,body,[class*="css"]{font-family:'Inter',sans-serif;background:#fff !important;}
#MainMenu,footer,header{visibility:hidden;}
section[data-testid="stSidebar"]{display:none !important;}
.block-container,[data-testid="stAppViewBlockContainer"],[data-testid="block-container"]{
    padding:0 60px 60px 60px !important;
    max-width:1560px !important;margin:0 auto !important;background:#fff !important;}
[data-testid="stAppViewBlockContainer"]>div:first-child{padding-top:0 !important;}

/* ── NAVBAR ── */
.top-header{background:#fff;border-bottom:2px solid #1B3A6B;padding:0 60px;height:64px;
    display:flex;align-items:center;justify-content:space-between;
    position:fixed;width:100%;top:0;z-index:999;
    box-shadow:0 2px 12px rgba(27,58,107,.08);margin:0 -60px 0 -60px;}
.header-brand{font-family:'Poppins',sans-serif;font-size:1.2rem;font-weight:700;color:#1B3A6B;}
.header-brand span{color:#E8A020;}
.header-nav{display:flex;gap:8px;align-items:center;}
.h-btn{background:#1B3A6B;color:#fff !important;border:none;padding:7px 18px;
    border-radius:6px;font-size:.82rem;font-weight:600;
    text-decoration:none !important;display:inline-block;}
.h-btn:hover{background:#243f7a;}
.h-btn-ghost{background:transparent;color:#1B3A6B !important;
    border:1.5px solid #1B3A6B;padding:6px 16px;border-radius:6px;
    font-size:.82rem;font-weight:600;text-decoration:none !important;display:inline-block;}
.h-btn-ghost:hover{background:#f0f4fb;}

/* ── PAGE TITLE ── */
.page-title-bar{background:#f7f9fc;border-left:4px solid #1B3A6B;
    padding:14px 60px;margin:40px -60px 28px -60px;}
.page-title-bar h2{font-family:'Poppins',sans-serif;font-size:1.15rem;
    font-weight:700;color:#1B3A6B;margin:0;}
.page-title-bar p{color:#6B7A8D;font-size:.8rem;margin:3px 0 0 0;}

/* ── FILTER ── */
.filter-wrap{background:#f7f9fc;border:1px solid #dde3ee;border-radius:10px;
    padding:22px 28px;margin-bottom:26px;}
.filter-label{font-size:.75rem;font-weight:600;color:#1B3A6B;
    letter-spacing:.8px;text-transform:uppercase;margin-bottom:6px;}
.stSelectbox>div>div{background:#fff !important;border:1.5px solid #c8d0e0 !important;
    border-radius:7px !important;color:#1B3A6B !important;font-size:.88rem !important;}
div[data-testid="stButton"]>button{background:#1B3A6B !important;color:#fff !important;
    font-family:'Montserrat',sans-serif !important;font-weight:700 !important;
    font-size:.9rem !important;letter-spacing:1px !important;border:none !important;
    border-radius:7px !important;padding:12px 0 !important;width:100% !important;}
div[data-testid="stButton"]>button:hover{background:#243f7a !important;}
div[data-testid="stDownloadButton"]>button{background:#1E7A34 !important;color:#fff !important;
    font-family:'Montserrat',sans-serif !important;font-weight:700 !important;
    font-size:.88rem !important;border:none !important;border-radius:7px !important;
    padding:11px 0 !important;width:100% !important;}

/* ── ROLE BADGE ── */
.role-badge{display:inline-block;padding:3px 10px;border-radius:12px;
    font-size:.7rem;font-weight:700;letter-spacing:.5px;text-transform:uppercase;
    margin-left:8px;vertical-align:middle;}
.role-bat{background:#EBF3FF;color:#1B3A6B;}
.role-bowl{background:#fff3e0;color:#b35a00;}
.role-all{background:#e8f5e9;color:#1E7A34;}

/* ── REPORT WRAPPER ── */
.rpt-wrap{background:#fff;border:2px solid #1B3A6B;border-radius:14px;
    padding:50px 60px;box-shadow:0 8px 48px rgba(27,58,107,.12);
    max-width:1020px;margin:0 auto 32px auto;}

/* ── REPORT MASTHEAD ── */
.rpt-masthead{background:linear-gradient(135deg,#1B3A6B 0%,#22488a 100%);
    border-radius:10px;padding:26px 32px;margin-bottom:36px;
    display:flex;align-items:center;justify-content:space-between;}
.rpt-masthead-left .rt{font-family:'Poppins',sans-serif;font-size:1.3rem;
    font-weight:800;color:#fff;letter-spacing:.3px;}
.rpt-masthead-left .rs{font-size:.8rem;color:rgba(255,255,255,.6);margin-top:4px;}
.rpt-meta-box{text-align:right;}
.rpt-meta-box .rdate{font-size:.75rem;color:rgba(255,255,255,.5);}
.rpt-meta-box .rid{display:inline-block;margin-top:6px;background:rgba(255,255,255,.12);
    border-radius:20px;padding:2px 12px;font-size:.7rem;font-weight:600;
    color:rgba(255,255,255,.75);letter-spacing:.4px;}

/* ── SUMMARY BOX ── */
.rpt-summary{background:#EBF3FF;border:1px solid #c0d0eb;border-radius:10px;
    padding:18px 24px;margin-bottom:28px;
    display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
.sum-item .sl{font-size:.7rem;color:#4A6FA5;font-weight:600;
    text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px;}
.sum-item .sv{font-family:'Poppins',sans-serif;font-size:.95rem;
    font-weight:700;color:#1B3A6B;}

/* ── PAGE DIVIDER ── */
.rpt-page-break{background:#f7f9fc;border:1px solid #dde3ee;border-radius:8px;
    padding:10px 18px;margin:32px 0 24px 0;
    display:flex;align-items:center;gap:10px;}
.rpt-page-num{background:#1B3A6B;color:#fff;padding:3px 12px;
    border-radius:20px;font-size:.72rem;font-weight:700;letter-spacing:.4px;}
.rpt-page-title{font-family:'Poppins',sans-serif;font-size:.9rem;
    font-weight:700;color:#1B3A6B;}

/* ── SECTION HEADING ── */
.rpt-sec{font-family:'Poppins',sans-serif;font-size:.95rem;font-weight:700;
    color:#1B3A6B;padding:10px 16px;background:#EBF3FF;
    border-left:5px solid #1B3A6B;border-radius:0 8px 8px 0;
    margin:24px 0 16px 0;display:flex;align-items:center;gap:8px;}
.rpt-sec-num{width:24px;height:24px;border-radius:50%;background:#1B3A6B;
    color:#fff;display:flex;align-items:center;justify-content:center;
    font-size:.72rem;font-weight:800;flex-shrink:0;}

/* ── COMPARISON GRID ── */
.cmp-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}
.cmp-col{background:#fff;border:1px solid #dde3ee;border-radius:10px;
    padding:16px 18px;position:relative;overflow:hidden;}
.cmp-col.p1-col::before{content:'';position:absolute;top:0;left:0;right:0;
    height:3px;background:#1B3A6B;}
.cmp-col.p2-col::before{content:'';position:absolute;top:0;left:0;right:0;
    height:3px;background:#E8A020;}
.cmp-player-name{font-family:'Poppins',sans-serif;font-size:.88rem;font-weight:700;
    color:#1B3A6B;margin-bottom:12px;padding-bottom:8px;
    border-bottom:1px solid #f0f3f8;}
.cmp-player-name.p2-name{color:#C07010;}
.cmp-stat{display:flex;justify-content:space-between;align-items:center;
    padding:7px 0;border-bottom:1px solid #f7f9fc;font-size:.83rem;}
.cmp-stat:last-child{border-bottom:none;}
.cmp-stat-lbl{color:#6B7A8D;font-weight:500;}
.cmp-stat-val{font-family:'Poppins',sans-serif;font-weight:700;color:#1B3A6B;font-size:.9rem;}
.cmp-stat-val.best{color:#1E7A34;}
.cmp-stat-val.worst{color:#C0392B;}

/* ── INSIGHT BOX ── */
.insight-box{background:#f7f9fc;border:1px solid #dde3ee;
    border-left:4px solid #E8A020;border-radius:0 8px 8px 0;
    padding:12px 18px;margin:12px 0;}
.ins-title{font-family:'Poppins',sans-serif;font-size:.72rem;font-weight:700;
    color:#8A6000;text-transform:uppercase;letter-spacing:.6px;margin-bottom:7px;}
.ins-item{display:flex;align-items:flex-start;gap:10px;
    padding:5px 0;font-size:.83rem;color:#2c3e50;line-height:1.6;}
.ins-item.pos{color:#1E7A34;}
.ins-item.neg{color:#C0392B;}
.ins-dot{width:6px;height:6px;border-radius:50%;background:#E8A020;
    flex-shrink:0;margin-top:8px;}

/* ── STRENGTH / WEAKNESS ── */
.sw-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}
.sw-box{border-radius:10px;padding:16px 18px;}
.sw-box.strength{background:#f0f9f2;border:1px solid #a8d4b0;}
.sw-box.weakness{background:#fdf0f0;border:1px solid #f0b8b8;}
.sw-box-title{font-family:'Poppins',sans-serif;font-size:.78rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.6px;margin-bottom:10px;}
.sw-box.strength .sw-box-title{color:#1E7A34;}
.sw-box.weakness .sw-box-title{color:#C0392B;}
.sw-item{display:flex;align-items:flex-start;gap:8px;
    font-size:.82rem;color:#2c3e50;padding:5px 0;line-height:1.5;}
.sw-item::before{content:"";flex-shrink:0;}
.sw-icon{flex-shrink:0;font-size:.85rem;margin-top:1px;}

/* ── VERDICT BOX ── */
.verdict-box{background:linear-gradient(135deg,#1B3A6B 0%,#243f7a 100%);
    border-radius:12px;padding:24px 28px;margin-top:20px;}
.verdict-title{font-family:'Poppins',sans-serif;font-size:.9rem;font-weight:800;
    color:#E8A020;text-transform:uppercase;letter-spacing:.8px;
    margin-bottom:14px;}
.verdict-text{font-size:.86rem;color:rgba(255,255,255,.9);line-height:1.8;}
.verdict-text strong{color:#E8A020;}

/* ── KEY INSIGHTS ── */
.key-insight-item{display:flex;align-items:flex-start;gap:12px;
    padding:11px 16px;background:#f7f9fc;border-radius:8px;
    border-left:4px solid #1B3A6B;margin-bottom:8px;
    font-size:.85rem;color:#2c3e50;line-height:1.6;}
.ki-num{width:22px;height:22px;border-radius:50%;background:#1B3A6B;
    color:#fff;font-size:.7rem;font-weight:800;display:flex;
    align-items:center;justify-content:center;flex-shrink:0;}

/* ── REPORT FOOTER ── */
.rpt-foot{background:#1B3A6B;border-radius:8px;padding:16px 24px;
    margin-top:36px;display:flex;align-items:center;justify-content:space-between;}
.rpt-foot-brand{font-family:'Poppins',sans-serif;font-size:.88rem;
    font-weight:700;color:#fff;}
.rpt-foot-brand span{color:#E8A020;}
.rpt-foot-meta{font-size:.73rem;color:rgba(255,255,255,.5);text-align:right;}

/* ── HR ── */
.hr{border:none;height:1px;background:#e8edf5;margin:22px 0;}

/* ── DL BAR ── */
.dl-bar{background:#f7f9fc;border-top:1px solid #dde3ee;padding:22px 60px;
    margin:0 -60px;display:flex;align-items:center;
    justify-content:space-between;flex-wrap:wrap;gap:14px;}
.dl-info{font-size:.82rem;color:#6B7A8D;line-height:1.5;}
.dl-info strong{color:#1B3A6B;}

/* ── NO DATA / FOOTER ── */
.no-data{text-align:center;padding:60px 20px;color:#8899aa;}
.no-data .icon{font-size:3rem;display:block;margin-bottom:12px;}
.site-footer{background:#1B3A6B;color:rgba(255,255,255,.7);text-align:center;
    padding:18px 60px;font-size:.8rem;margin:0 -60px -60px -60px;letter-spacing:.3px;}
.site-footer span{color:#E8A020;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  DATA LOADING
# ─────────────────────────────────────────────
NON_NUM = {"Match_id","Match_date","Opponent","Venue","Country","Match_result",
           "Dismissed_method","Dismissed_Length","Dismissed_against",
           "Dismissed_by","Dismissed _by"}

@st.cache_data
def load_data():
    master = pd.read_csv(MASTER_PATH)
    master.columns = master.columns.str.strip()
    wb    = openpyxl.load_workbook(STATS_PATH)
    stats = {}
    for sh in wb.sheetnames:
        ws   = wb[sh]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            stats[sh.strip()] = pd.DataFrame(); continue
        cols = [str(c).strip() if c else f"_c{i}" for i,c in enumerate(rows[0])]
        df   = pd.DataFrame(rows[1:], columns=cols)
        df["Match_date"] = pd.to_datetime(df["Match_date"], errors="coerce")
        for c in df.columns:
            if c not in NON_NUM:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        stats[sh.strip()] = df
    return master, stats

try:
    master_df, stats_dict = load_data()
except Exception as e:
    st.error(f"Data load error: {e}"); st.stop()

ALL_PLAYERS = sorted(stats_dict.keys())

# ─────────────────────────────────────────────
#  ROLE HELPERS
# ─────────────────────────────────────────────
def get_role(name):
    """Return normalised role string: 'Batsman', 'Bowler', or 'All-Rounder'."""
    row = master_df[master_df["PlayerName"].str.strip() == name]
    if not row.empty:
        return str(row.iloc[0]["Role"]).strip()
    df = stats_dict.get(name, pd.DataFrame())
    has_bat  = "Batted_Flag" in df.columns
    has_bowl = "Bowl_Flag"   in df.columns
    if has_bat and has_bowl: return "All-Rounder"
    if has_bowl:             return "Bowler"
    return "Batsman"

def role_norm(role_str):
    """Normalised lowercase role for logic checks."""
    r = role_str.lower()
    if "all" in r or "round" in r: return "allrounder"
    if "bowl" in r:                return "bowler"
    return "batsman"

def get_meta(name):
    row = master_df[master_df["PlayerName"].str.strip() == name]
    if not row.empty:
        p = row.iloc[0]
        return {k: str(p.get(k,"—")) for k in
                ["Country","Role","BattingStyle","BowlingStyle","DOB","DebutYear"]}
    return {"Country":"India","Role":get_role(name),"BattingStyle":"RHB",
            "BowlingStyle":"—","DOB":"N/A","DebutYear":"N/A"}

def bat_rows(df):
    return df[df["Batted_Flag"]==1] if "Batted_Flag" in df.columns else pd.DataFrame()

def bowl_rows(df):
    return df[df["Bowl_Flag"]==1] if "Bowl_Flag" in df.columns else pd.DataFrame()

def gcol(df, *names):
    for n in names:
        if n in df.columns: return df[n]
    return pd.Series(0, index=df.index)

def calc_avg(runs_s, not_out_s, inn):
    tot = float(runs_s.sum())
    dis = max(int(inn) - int(not_out_s.sum()), 0)
    return round(tot/dis, 2) if dis > 0 else round(tot, 2)

def safe_r(v, decimals=2):
    try: return round(float(v), decimals)
    except: return 0.0

def strip_html(t):
    return re.sub(r"<[^>]+>","",str(t)).replace("&amp;","&").replace("&nbsp;"," ")

def apply_filters(df, year_filter):
    if year_filter != "Overall":
        df = df[df["Match_date"].dt.year == int(year_filter)]
    return df

# ─────────────────────────────────────────────
#  STATS CALCULATORS
# ─────────────────────────────────────────────
def bat_stats(bat_df):
    if bat_df.empty:
        return dict(inn=0, runs=0, avg=0.0, sr=0.0, f50=0, f100=0,
                    f4=0, f6=0, dots=0, best="0", win_pct=0.0, no=0)
    inn  = len(bat_df)
    runs = int(bat_df["Runs"].sum())
    bf   = int(bat_df["Balls_faced"].sum()) if "Balls_faced" in bat_df.columns else 0
    no_s = bat_df.get("Not_out", pd.Series(0, index=bat_df.index))
    no   = int(no_s.sum())
    avg  = calc_avg(bat_df["Runs"], no_s, inn)
    sr   = safe_r(bat_df["Strike_Rate"].replace(0, float("nan")).mean(), 1)
    f50  = int(bat_df["50s"].sum()) if "50s" in bat_df.columns else 0
    f100 = int(bat_df["100s"].sum()) if "100s" in bat_df.columns else 0
    f4   = int(bat_df["4s"].sum())
    f6   = int(bat_df["6s"].sum())
    dots = int(bat_df["Dots"].sum()) if "Dots" in bat_df.columns else 0
    if not bat_df.empty:
        bi   = bat_df["Runs"].idxmax()
        best = f"{int(bat_df.loc[bi,'Runs'])}{'*' if int(bat_df.loc[bi,'Not_out']) else ''}"
    else: best = "0"
    wp  = round((bat_df["Match_result"]=="Won").sum() / max(inn,1) * 100, 1) if "Match_result" in bat_df.columns else 0.0
    return dict(inn=inn, runs=runs, avg=avg, sr=sr, f50=f50, f100=f100,
                f4=f4, f6=f6, dots=dots, best=best, win_pct=wp, no=no)

def bowl_stats(bowl_df):
    """
    Returns bowling stats including bowling_sr = Balls_Bowled / Wickets
    (lower is better — fewer balls needed per wicket).
    """
    if bowl_df.empty:
        return dict(inn=0, wkts=0, eco=0.0, avg=0.0, bowling_sr=0.0,
                    dots=0, dot_pct=0.0, best="0/0", mai=0,
                    balls=0, runs_con=0)
    inn       = len(bowl_df)
    tbl       = int(bowl_df["Balls_Bowled"].sum()) if "Balls_Bowled" in bowl_df.columns else 0
    wkts      = int(bowl_df["Wickets"].sum())
    rc        = int(bowl_df["Runs_Conceded"].sum()) if "Runs_Conceded" in bowl_df.columns else 0
    eco       = safe_r(bowl_df["Economy"].replace(0, float("nan")).mean(), 2)
    ba        = round(rc / max(wkts, 1), 2)
    # ── Bowling Strike Rate: balls bowled per wicket (lower = better) ──────
    bowling_sr = round(tbl / max(wkts, 1), 1)
    dots      = int(gcol(bowl_df,"Dot_Balls").sum())
    dp        = round(dots / max(tbl,1) * 100, 1)
    mai       = int(bowl_df["Maidens"].sum()) if "Maidens" in bowl_df.columns else 0
    if wkts > 0:
        br   = bowl_df.sort_values(["Wickets","Runs_Conceded"],ascending=[False,True]).iloc[0]
        best = f"{int(br['Wickets'])}/{int(br.get('Runs_Conceded',0))}"
    else: best = "0/0"
    return dict(inn=inn, wkts=wkts, eco=eco, avg=ba, bowling_sr=bowling_sr,
                dots=dots, dot_pct=dp, best=best, mai=mai,
                balls=tbl, runs_con=rc)

def best_opponent(bat_df):
    if bat_df.empty or "Opponent" not in bat_df.columns: return "—", 0.0
    og = bat_df.groupby("Opponent").apply(
        lambda g: calc_avg(g["Runs"],
                           g.get("Not_out",pd.Series(0,index=g.index)), len(g)))
    if og.empty: return "—", 0.0
    bo = og.idxmax()
    return str(bo), safe_r(og.max(), 2)

def best_opponent_bowl(bowl_df):
    """Best bowling opponent = most wickets taken against."""
    if bowl_df.empty or "Opponent" not in bowl_df.columns: return "—", 0
    og = bowl_df.groupby("Opponent")["Wickets"].sum()
    if og.empty: return "—", 0
    bo = og.idxmax()
    return str(bo), int(og.max())

def best_venue(bat_df):
    if bat_df.empty or "Venue" not in bat_df.columns: return "—", 0.0
    vg = bat_df.groupby("Venue").apply(
        lambda g: calc_avg(g["Runs"],
                           g.get("Not_out",pd.Series(0,index=g.index)), len(g)))
    if vg.empty: return "—", 0.0
    bv = vg.idxmax()
    return str(bv), safe_r(vg.max(), 2)

def best_venue_bowl(bowl_df):
    """Best bowling venue = most wickets."""
    if bowl_df.empty or "Venue" not in bowl_df.columns: return "—", 0
    vg = bowl_df.groupby("Venue")["Wickets"].sum()
    if vg.empty: return "—", 0
    bv = vg.idxmax()
    return str(bv), int(vg.max())

def recent_form(bat_df, bowl_df, n=5):
    r_bat  = bat_df.sort_values("Match_date").tail(n) if not bat_df.empty else pd.DataFrame()
    r_bowl = bowl_df.sort_values("Match_date").tail(n) if not bowl_df.empty else pd.DataFrame()
    return bat_stats(r_bat), bowl_stats(r_bowl)

# ─────────────────────────────────────────────
#  ROLE-AWARE METRIC BUILDERS
# ─────────────────────────────────────────────
def build_bat_metrics(bsa, bsb):
    """Full batting comparison metrics list."""
    return [
        ("Innings",          bsa["inn"],  bsb["inn"],  True),
        ("Total Runs",       bsa["runs"], bsb["runs"], True),
        ("Batting Average",  bsa["avg"],  bsb["avg"],  True),
        ("Strike Rate",      bsa["sr"],   bsb["sr"],   True),
        ("Best Score",       bsa["best"], bsb["best"], True),
        ("50s",              bsa["f50"],  bsb["f50"],  True),
        ("100s",             bsa["f100"], bsb["f100"], True),
        ("4s",               bsa["f4"],   bsb["f4"],   True),
        ("6s",               bsa["f6"],   bsb["f6"],   True),
        ("Not Outs",         bsa["no"],   bsb["no"],   True),
        ("Win %",            f"{bsa['win_pct']}%", f"{bsb['win_pct']}%", True),
    ]

def build_bowl_metrics(bwa, bwb):
    """
    Full bowling comparison metrics.
    Bowling Strike Rate = Balls Bowled / Wickets  (lower is better).
    """
    return [
        ("Innings",             bwa["inn"],          bwb["inn"],          True),
        ("Total Wickets",       bwa["wkts"],         bwb["wkts"],         True),
        ("Bowling Average",     bwa["avg"],          bwb["avg"],          False),
        ("Economy Rate",        bwa["eco"],          bwb["eco"],          False),
        ("Bowling Strike Rate", bwa["bowling_sr"],   bwb["bowling_sr"],   False),  # lower = better
        ("Maidens",             bwa["mai"],          bwb["mai"],          True),
        ("Dot Ball %",          f"{bwa['dot_pct']}%",f"{bwb['dot_pct']}%",True),
        ("Best Figures",        bwa["best"],         bwb["best"],         True),
    ]

def build_form_metrics(fb_a_bat, fb_b_bat, fb_a_bowl, fb_b_bowl, player_role):
    """Recent-form metrics tailored to role."""
    rn = role_norm(player_role)
    metrics = []
    if rn in ("batsman", "allrounder"):
        metrics += [
            ("Runs (Last 5)",    fb_a_bat["runs"],  fb_b_bat["runs"],  True),
            ("Batting Average",  fb_a_bat["avg"],   fb_b_bat["avg"],   True),
            ("Strike Rate",      fb_a_bat["sr"],    fb_b_bat["sr"],    True),
            ("50s",              fb_a_bat["f50"],   fb_b_bat["f50"],   True),
        ]
    if rn in ("bowler", "allrounder"):
        metrics += [
            ("Wickets (Last 5)",     fb_a_bowl["wkts"],       fb_b_bowl["wkts"],       True),
            ("Economy",              fb_a_bowl["eco"],        fb_b_bowl["eco"],        False),
            ("Bowling SR (Last 5)",  fb_a_bowl["bowling_sr"], fb_b_bowl["bowling_sr"], False),
        ]
    return metrics

# ─────────────────────────────────────────────
#  INSIGHT GENERATORS  (role-aware)
# ─────────────────────────────────────────────
def gen_bat_insights(bs1, bs2, n1, n2):
    ins = []
    if bs1["avg"] > bs2["avg"]:
        ins.append(f"<strong>{n1}</strong> shows better batting consistency "
                   f"(Avg {bs1['avg']} vs {bs2['avg']}).")
    elif bs2["avg"] > bs1["avg"]:
        ins.append(f"<strong>{n2}</strong> shows better batting consistency "
                   f"(Avg {bs2['avg']} vs {bs1['avg']}).")
    if bs1["sr"] > bs2["sr"]:
        ins.append(f"<strong>{n1}</strong> plays a more aggressive role "
                   f"(SR {bs1['sr']} vs {bs2['sr']}).")
    elif bs2["sr"] > bs1["sr"]:
        ins.append(f"<strong>{n2}</strong> plays a more aggressive role "
                   f"(SR {bs2['sr']} vs {bs1['sr']}).")
    if bs1["runs"] > bs2["runs"]:
        ins.append(f"<strong>{n1}</strong> has contributed more runs overall "
                   f"({bs1['runs']} vs {bs2['runs']}).")
    elif bs2["runs"] > bs1["runs"]:
        ins.append(f"<strong>{n2}</strong> has contributed more runs overall "
                   f"({bs2['runs']} vs {bs1['runs']}).")
    return ins

def gen_bowl_insights(bw1, bw2, n1, n2):
    ins = []
    # Economy
    if bw1["eco"] and bw2["eco"] and bw1["eco"] < bw2["eco"]:
        ins.append(f"<strong>{n1}</strong> demonstrates better run control "
                   f"(Economy {bw1['eco']} vs {bw2['eco']}).")
    elif bw2["eco"] and bw1["eco"] and bw2["eco"] < bw1["eco"]:
        ins.append(f"<strong>{n2}</strong> demonstrates better run control "
                   f"(Economy {bw2['eco']} vs {bw1['eco']}).")
    # Wickets
    if bw1["wkts"] > bw2["wkts"]:
        ins.append(f"<strong>{n1}</strong> has stronger wicket-taking ability "
                   f"({bw1['wkts']} vs {bw2['wkts']} wickets).")
    elif bw2["wkts"] > bw1["wkts"]:
        ins.append(f"<strong>{n2}</strong> has stronger wicket-taking ability "
                   f"({bw2['wkts']} vs {bw1['wkts']} wickets).")
    # Bowling Strike Rate (lower = better)
    if bw1["bowling_sr"] and bw2["bowling_sr"]:
        if bw1["bowling_sr"] < bw2["bowling_sr"]:
            ins.append(f"<strong>{n1}</strong> takes wickets faster "
                       f"(Bowling SR {bw1['bowling_sr']} vs {bw2['bowling_sr']} balls/wicket).")
        elif bw2["bowling_sr"] < bw1["bowling_sr"]:
            ins.append(f"<strong>{n2}</strong> takes wickets faster "
                       f"(Bowling SR {bw2['bowling_sr']} vs {bw1['bowling_sr']} balls/wicket).")
    return ins

def gen_form_insight(fb1_bat, fb2_bat, fb1_bowl, fb2_bowl, n1, n2, player_role):
    rn = role_norm(player_role)
    if rn == "bowler":
        if fb1_bowl["wkts"] > fb2_bowl["wkts"] and fb1_bowl["eco"] <= fb2_bowl["eco"]:
            return (f"<strong>{n1}</strong> is in significantly better recent bowling form — "
                    f"more wickets ({fb1_bowl['wkts']} vs {fb2_bowl['wkts']}) and lower economy "
                    f"({fb1_bowl['eco']} vs {fb2_bowl['eco']}) in the last 5 matches.")
        elif fb2_bowl["wkts"] > fb1_bowl["wkts"] and fb2_bowl["eco"] <= fb1_bowl["eco"]:
            return (f"<strong>{n2}</strong> is in significantly better recent bowling form — "
                    f"more wickets ({fb2_bowl['wkts']} vs {fb1_bowl['wkts']}) and lower economy "
                    f"({fb2_bowl['eco']} vs {fb1_bowl['eco']}) in the last 5 matches.")
        elif fb1_bowl["wkts"] > fb2_bowl["wkts"]:
            return (f"<strong>{n1}</strong> has taken more wickets recently "
                    f"({fb1_bowl['wkts']} vs {fb2_bowl['wkts']}), though economies are comparable.")
        else:
            return (f"Both bowlers show comparable recent form — "
                    f"{n1} SR {fb1_bowl['bowling_sr']} vs {n2} SR {fb2_bowl['bowling_sr']}.")
    # batsman / allrounder — use batting form
    fb1, fb2 = fb1_bat, fb2_bat
    if fb1["runs"] > fb2["runs"] and fb1["avg"] > fb2["avg"]:
        return (f"<strong>{n1}</strong> is in significantly better recent batting form — "
                f"more runs ({fb1['runs']} vs {fb2['runs']}) and higher average "
                f"({fb1['avg']} vs {fb2['avg']}) over the last 5 matches.")
    elif fb2["runs"] > fb1["runs"] and fb2["avg"] > fb1["avg"]:
        return (f"<strong>{n2}</strong> is in significantly better recent batting form — "
                f"more runs ({fb2['runs']} vs {fb1['runs']}) and higher average "
                f"({fb2['avg']} vs {fb1['avg']}) over the last 5 matches.")
    elif fb1["runs"] > fb2["runs"]:
        return (f"<strong>{n1}</strong> has scored more runs recently ({fb1['runs']} vs {fb2['runs']}), "
                f"though both players show comparable averages.")
    else:
        return (f"Both players show comparable recent batting form, with <strong>{n1}</strong> "
                f"averaging {fb1['avg']} and <strong>{n2}</strong> averaging {fb2['avg']} "
                f"in the last 5 matches.")

def gen_venue_insight(v1, a1, v2, a2, n1, n2, player_role):
    rn = role_norm(player_role)
    if rn == "bowler":
        if a1 >= a2:
            return (f"<strong>{n1}</strong> is most effective at {v1} ({a1} wickets), while "
                    f"<strong>{n2}</strong> performs best at {v2} ({a2} wickets). "
                    f"{n1} holds the venue edge.")
        return (f"<strong>{n2}</strong> leads venue performance ({a2} wickets at {v2}), "
                f"compared to <strong>{n1}</strong>'s {a1} wickets at {v1}.")
    if a1 >= a2:
        return (f"<strong>{n1}</strong> performs best at {v1} (avg {a1}), while "
                f"<strong>{n2}</strong> excels at {v2} (avg {a2}). "
                f"{n1} holds the edge in venue-specific batting.")
    return (f"<strong>{n2}</strong> leads venue batting (avg {a2} at {v2}), "
            f"compared to <strong>{n1}</strong>'s best average of {a1} at {v1}.")

def gen_opp_insight(o1, a1, o2, a2, n1, n2, player_role):
    rn = role_norm(player_role)
    if rn == "bowler":
        if a1 >= a2:
            return (f"<strong>{n1}</strong> dominates against {o1} ({a1} wickets), while "
                    f"<strong>{n2}</strong> performs best against {o2} ({a2} wickets).")
        return (f"<strong>{n2}</strong> shows greater opponent dominance ({a2} wickets vs {o2}), "
                f"compared to <strong>{n1}</strong>'s {a1} against {o1}.")
    if a1 >= a2:
        return (f"<strong>{n1}</strong> dominates against {o1} (avg {a1}), while "
                f"<strong>{n2}</strong> performs best against {o2} (avg {a2}).")
    return (f"<strong>{n2}</strong> shows greater opponent dominance (avg {a2} vs {o2}), "
            f"compared to <strong>{n1}</strong>'s {a1} against {o1}.")

def gen_strengths(bs, bw, name, rn):
    s = []
    if rn in ("batsman", "allrounder"):
        if bs["avg"] >= 40:  s.append(f"High batting consistency (avg {bs['avg']})")
        if bs["sr"] >= 100:  s.append(f"Aggressive scoring style (SR {bs['sr']})")
        if bs["f100"] >= 1:  s.append(f"Capable of match-defining centuries ({bs['f100']} tons)")
        elif bs["f50"] >= 2: s.append(f"Regular half-century scorer ({bs['f50']} fifties)")
        bnd = bs["f4"]*4 + bs["f6"]*6
        if bs["runs"] > 0 and bnd/bs["runs"] >= 0.45:
            s.append("Strong boundary scoring ability")
        if bs["win_pct"] >= 60: s.append(f"Wins more often when batting well ({bs['win_pct']}%)")
    if rn in ("bowler", "allrounder"):
        if bw["eco"] > 0 and bw["eco"] <= 5.5:
            s.append(f"Excellent economy rate ({bw['eco']} RPO)")
        if bw["wkts"] >= 5:
            s.append(f"Match-winning bowler ({bw['wkts']} wickets)")
        if bw["bowling_sr"] > 0 and bw["bowling_sr"] <= 20:
            s.append(f"Rapid wicket-taker (Bowling SR {bw['bowling_sr']})")
        if bw["dot_pct"] >= 40:
            s.append(f"High dot ball pressure ({bw['dot_pct']}% dots)")
        if bw["mai"] >= 3:
            s.append(f"Effective in containing runs ({bw['mai']} maidens)")
    return s if s else ["Consistent match contributor"]

def gen_weaknesses(bs, bw, name, rn):
    w = []
    if rn in ("batsman", "allrounder"):
        if bs["avg"] < 30 and bs["inn"] > 0:
            w.append(f"Lower batting average ({bs['avg']}) — needs improvement")
        if bs["sr"] < 80 and bs["inn"] > 0:
            w.append(f"Below-par strike rate ({bs['sr']}) — scoring rate concern")
        if bs["f50"] == 0 and bs["f100"] == 0 and bs["inn"] >= 3:
            w.append("Struggles to convert starts into milestone scores")
    if rn in ("bowler", "allrounder"):
        if bw["eco"] > 7:
            w.append(f"High economy rate ({bw['eco']}) — leaks runs")
        if bw["wkts"] == 0 and bw["inn"] > 3:
            w.append("Low wicket-taking impact across recent matches")
        if bw["bowling_sr"] > 35 and bw["inn"] > 3:
            w.append(f"Slow to take wickets (Bowling SR {bw['bowling_sr']} balls/wicket)")
    return w if w else ["Continues to develop all aspects of game"]

def gen_key_insights(bs1, bs2, bw1, bw2, fb1_bat, fb2_bat, fb1_bowl, fb2_bowl,
                     n1, n2, rn):
    ki = []
    if rn in ("batsman", "allrounder"):
        if bs1["avg"] > bs2["avg"]:
            ki.append(f"{n1} shows higher batting consistency with an average of "
                      f"{bs1['avg']} compared to {n2}'s {bs2['avg']}.")
        elif bs2["avg"] > bs1["avg"]:
            ki.append(f"{n2} shows higher batting consistency with an average of "
                      f"{bs2['avg']} compared to {n1}'s {bs1['avg']}.")
        if abs(bs1["sr"] - bs2["sr"]) > 5:
            aggr = n1 if bs1["sr"] > bs2["sr"] else n2
            ki.append(f"{aggr} demonstrates a more aggressive batting style based on strike rate.")
    if rn in ("bowler", "allrounder"):
        if bw1["wkts"] != bw2["wkts"]:
            bbetter = n1 if bw1["wkts"] > bw2["wkts"] else n2
            ki.append(f"{bbetter} contributes more with the ball in terms of wickets taken.")
        if bw1["eco"] and bw2["eco"] and abs(bw1["eco"] - bw2["eco"]) >= 0.5:
            ec = n1 if bw1["eco"] < bw2["eco"] else n2
            ki.append(f"{ec} offers better bowling economy, making them harder to score against.")
        # Bowling SR comparison
        if bw1["bowling_sr"] and bw2["bowling_sr"] and bw1["bowling_sr"] != bw2["bowling_sr"]:
            faster = n1 if bw1["bowling_sr"] < bw2["bowling_sr"] else n2
            ki.append(f"{faster} takes wickets quicker — lower Bowling Strike Rate "
                      f"({min(bw1['bowling_sr'],bw2['bowling_sr'])} balls/wicket).")
    # Recent form
    if rn == "bowler":
        if fb1_bowl["wkts"] != fb2_bowl["wkts"]:
            fr = n1 if fb1_bowl["wkts"] > fb2_bowl["wkts"] else n2
            ki.append(f"{fr} is in superior recent bowling form based on last 5 matches.")
    else:
        if fb1_bat["runs"] != fb2_bat["runs"]:
            fr = n1 if fb1_bat["runs"] > fb2_bat["runs"] else n2
            ki.append(f"{fr} is in superior recent batting form based on last 5 matches.")
    if rn in ("batsman", "allrounder") and bs1["runs"] > 0 and bs2["runs"] > 0:
        rn2 = n1 if bs1["runs"] > bs2["runs"] else n2
        ki.append(f"{rn2} has made a greater overall batting contribution across the analysis period.")
    return ki[:5]

def gen_verdict(bs1, bs2, bw1, bw2, n1, n2, rn):
    if rn == "batsman":
        consistent = n1 if bs1["avg"] >= bs2["avg"] else n2
        aggressive  = n1 if bs1["sr"]  >= bs2["sr"]  else n2
        v  = (f"<strong>{consistent}</strong> emerges as the more consistent batter "
              f"with a higher average across the analysis period. ")
        v += (f"<strong>{aggressive}</strong> brings greater match impact through an "
              f"aggressive scoring approach. ")
        v += (f"Both players offer distinct strengths — while <strong>{consistent}</strong> "
              f"anchors the innings, <strong>{aggressive}</strong> provides the firepower "
              f"to accelerate the scoring rate.")
        return v
    if rn == "bowler":
        wicket_winner = n1 if bw1["wkts"] >= bw2["wkts"] else n2
        economical    = n1 if (bw1["eco"] or 99) <= (bw2["eco"] or 99) else n2
        faster_sr     = n1 if (bw1["bowling_sr"] or 99) <= (bw2["bowling_sr"] or 99) else n2
        v  = (f"<strong>{wicket_winner}</strong> emerges as the superior wicket-taker "
              f"with more wickets across the analysis period. ")
        v += (f"<strong>{economical}</strong> is the more economical option, "
              f"restricting the opposition effectively. ")
        v += (f"In terms of strike rate, <strong>{faster_sr}</strong> needs fewer balls "
              f"per wicket — a critical advantage in T20 and pressure situations. ")
        v += (f"Both bowlers bring distinct strengths to the attack.")
        return v
    # allrounder
    consistent = n1 if bs1["avg"] >= bs2["avg"] else n2
    bowler_best = n1 if bw1["wkts"] >= bw2["wkts"] else n2
    v  = (f"<strong>{consistent}</strong> leads with the bat showing a higher average. ")
    v += (f"In the bowling department, <strong>{bowler_best}</strong> provides "
          f"the stronger wicket-taking threat. ")
    v += (f"Both all-rounders offer a balanced contribution — the choice between them "
          f"depends on whether the team needs more batting solidity or bowling penetration.")
    return v

# ─────────────────────────────────────────────
#  HTML BUILDERS
# ─────────────────────────────────────────────
def sec(num, icon, title):
    return (f'<div class="rpt-sec">'
            f'<span class="rpt-sec-num">{num}</span>'
            f'{icon} &nbsp;{title}</div>')

def cmp_row(lbl, v1, v2, higher_better=True):
    try:
        n1 = float(str(v1).replace("*","").replace("%",""))
        n2 = float(str(v2).replace("*","").replace("%",""))
        if n1 == n2:
            c1 = c2 = ""
        elif (higher_better and n1 > n2) or (not higher_better and n1 < n2):
            c1, c2 = "best", "worst"
        else:
            c1, c2 = "worst", "best"
    except:
        c1 = c2 = ""
    return (f'<div class="cmp-stat">'
            f'<span class="cmp-stat-lbl">{lbl}</span>'
            f'<span class="cmp-stat-val {c1}">{v1}</span>'
            f'</div>',
            f'<div class="cmp-stat">'
            f'<span class="cmp-stat-lbl">{lbl}</span>'
            f'<span class="cmp-stat-val {c2}">{v2}</span>'
            f'</div>')

def insight_block(items):
    if not items: return ""
    rows = "".join(f'<div class="ins-item"><span class="ins-dot"></span>{i}</div>'
                   for i in items)
    return f'<div class="insight-box"><div class="ins-title">Insights</div>{rows}</div>'

def single_insight(text):
    return (f'<div class="insight-box">'
            f'<div class="ins-title">Insight</div>'
            f'<div class="ins-item"><span class="ins-dot"></span>{text}</div>'
            f'</div>')

def sw_item(text, icon="•"):
    return f'<div class="sw-item"><span class="sw-icon">{icon}</span>{text}</div>'

def key_insight_item(num, text):
    return (f'<div class="key-insight-item">'
            f'<span class="ki-num">{num}</span>'
            f'{text}</div>')

def page_break_div(num, title):
    return (f'<div class="rpt-page-break">'
            f'<span class="rpt-page-num">PAGE {num}</span>'
            f'<span class="rpt-page-title">{title}</span>'
            f'</div>')

def role_badge_html(role):
    rn = role_norm(role)
    cls = {"batsman":"role-bat","bowler":"role-bowl","allrounder":"role-all"}.get(rn,"role-bat")
    return f'<span class="role-badge {cls}">{role}</span>'

# ─────────────────────────────────────────────
#  NAVBAR
# ─────────────────────────────────────────────
st.markdown("""
<div class="top-header">
  <div class="header-brand">Cricket Player<span> PerformanceAnalysis</span></div>
  <div class="header-nav">
    <a class="h-btn" href="/"         target="_self">Back</a>
    <a class="h-btn" href="/admin"    target="_self">Admin</a>
    <a class="h-btn" href="/userhome" target="_self">Home</a>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="page-title-bar">
  <h2>Head-to-Head Player Comparison Report</h2>
  <p>Role-aware comparison — Batsmen show batting metrics · Bowlers show bowling strike rate (balls/wicket) · All-Rounders show both</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  FILTER PANEL
# ─────────────────────────────────────────────
def_p1 = st.session_state.get("cmp_p1", ALL_PLAYERS[0])
def_p2 = st.session_state.get("cmp_p2",
         [p for p in ALL_PLAYERS if p != def_p1][0])
fc1, fc2, fc3, fc4 = st.columns([3, 3, 2, 2])

with fc1:
    st.markdown('<div class="filter-label">Player A</div>', unsafe_allow_html=True)
    p1_idx = ALL_PLAYERS.index(def_p1) if def_p1 in ALL_PLAYERS else 0
    sel_pa = st.selectbox("PA", ALL_PLAYERS, index=p1_idx, label_visibility="collapsed")

with fc2:
    st.markdown('<div class="filter-label">Player B</div>', unsafe_allow_html=True)
    p2_opts = [p for p in ALL_PLAYERS if p != sel_pa]
    def_p2  = def_p2 if def_p2 in p2_opts else p2_opts[0]
    sel_pb  = st.selectbox("PB", p2_opts, index=p2_opts.index(def_p2),
                           label_visibility="collapsed")

with fc3:
    st.markdown('<div class="filter-label">Role Focus</div>', unsafe_allow_html=True)
    ra = get_role(sel_pa); rb = get_role(sel_pb)
    rna = role_norm(ra);   rnb = role_norm(rb)
    role_opts = []
    if (rna in ("batsman","allrounder")) and (rnb in ("batsman","allrounder")):
        role_opts.append("Batting")
    if (rna in ("bowler","allrounder")) and (rnb in ("bowler","allrounder")):
        role_opts.append("Bowling")
    if not role_opts: role_opts = ["Batting"]
    if len(role_opts) > 1: role_opts.insert(0, "All-Round")
    sel_role = st.selectbox("Role", role_opts, label_visibility="collapsed")

with fc4:
    st.markdown('<div class="filter-label">Period</div>', unsafe_allow_html=True)
    _df1  = stats_dict.get(sel_pa, pd.DataFrame())
    _df2  = stats_dict.get(sel_pb, pd.DataFrame())
    _yrs  = sorted(set(
        list(_df1["Match_date"].dt.year.dropna().astype(int).unique() if not _df1.empty else []) +
        list(_df2["Match_date"].dt.year.dropna().astype(int).unique() if not _df2.empty else [])
    ))
    sel_yr = st.selectbox("Year", ["Overall"] + [str(y) for y in _yrs],
                          label_visibility="collapsed")

st.markdown('<div class="filter-label">&nbsp;</div>', unsafe_allow_html=True)
gen_btn = st.button("GENERATE", use_container_width=True)

# ─────────────────────────────────────────────
#  SESSION GATE
# ─────────────────────────────────────────────
if gen_btn:
    st.session_state.update(cr_ok=True, cr_pa=sel_pa, cr_pb=sel_pb,
                            cr_role=sel_role, cr_yr=sel_yr)

if not st.session_state.get("cr_ok"):
    st.markdown("""
    <div class="no-data">
      <span class="icon"></span>
      <p>Select two players and click <strong>Generate Report</strong></p>
    </div>""", unsafe_allow_html=True)
    st.stop()

sel_pa   = st.session_state["cr_pa"]
sel_pb   = st.session_state["cr_pb"]
sel_role = st.session_state["cr_role"]
sel_yr   = st.session_state["cr_yr"]

# ─────────────────────────────────────────────
#  BUILD DATA
# ─────────────────────────────────────────────
dfa = apply_filters(stats_dict.get(sel_pa, pd.DataFrame()), sel_yr)
dfb = apply_filters(stats_dict.get(sel_pb, pd.DataFrame()), sel_yr)

bata  = bat_rows(dfa);  batb  = bat_rows(dfb)
bowla = bowl_rows(dfa); bowlb = bowl_rows(dfb)

bsa  = bat_stats(bata);   bsb  = bat_stats(batb)
bwa  = bowl_stats(bowla); bwb  = bowl_stats(bowlb)

fb_a_bat, fb_a_bowl = recent_form(bata, bowla)
fb_b_bat, fb_b_bowl = recent_form(batb, bowlb)

# Role of each player (for venue/opponent context)
role_a = get_role(sel_pa)
role_b = get_role(sel_pb)
rna    = role_norm(role_a)
rnb    = role_norm(role_b)

# Determine combined role for the comparison session
rl_norm_combined = role_norm(sel_role)  # "batsman" | "bowler" | "allrounder"

# Venue & opponent — use bowling stats for bowlers, batting for batsmen/AR
if rna == "bowler":
    opp_a, opp_a_val = best_opponent_bowl(bowla)
    ven_a, ven_a_val = best_venue_bowl(bowla)
else:
    opp_a, opp_a_val = best_opponent(bata)
    ven_a, ven_a_val = best_venue(bata)

if rnb == "bowler":
    opp_b, opp_b_val = best_opponent_bowl(bowlb)
    ven_b, ven_b_val = best_venue_bowl(bowlb)
else:
    opp_b, opp_b_val = best_opponent(batb)
    ven_b, ven_b_val = best_venue(batb)

show_bat  = rl_norm_combined in ("batsman", "allrounder")
show_bowl = rl_norm_combined in ("bowler",  "allrounder")

matches_compared = max(len(dfa), len(dfb))
gen_ts   = datetime.now().strftime("%d %B %Y,  %H:%M")
gen_date = datetime.now().strftime("%d %B %Y")
rpt_id   = f"CMP-{datetime.now().strftime('%Y%m%d%H%M')}"

# Pre-compute insights
ins_bat  = gen_bat_insights(bsa, bsb, sel_pa, sel_pb)  if show_bat  else []
ins_bowl = gen_bowl_insights(bwa, bwb, sel_pa, sel_pb) if show_bowl else []
form_ins = gen_form_insight(fb_a_bat, fb_b_bat, fb_a_bowl, fb_b_bowl,
                            sel_pa, sel_pb, sel_role)
opp_ins  = gen_opp_insight(opp_a, opp_a_val, opp_b, opp_b_val,
                           sel_pa, sel_pb, sel_role)
ven_ins  = gen_venue_insight(ven_a, ven_a_val, ven_b, ven_b_val,
                             sel_pa, sel_pb, sel_role)

str_a = gen_strengths(bsa, bwa, sel_pa, rna)
str_b = gen_strengths(bsb, bwb, sel_pb, rnb)
wk_a  = gen_weaknesses(bsa, bwa, sel_pa, rna)
wk_b  = gen_weaknesses(bsb, bwb, sel_pb, rnb)

key_ins = gen_key_insights(bsa, bsb, bwa, bwb, fb_a_bat, fb_b_bat,
                           fb_a_bowl, fb_b_bowl, sel_pa, sel_pb,
                           rl_norm_combined)
verdict = gen_verdict(bsa, bsb, bwa, bwb, sel_pa, sel_pb, rl_norm_combined)

# ═══════════════════════════════════════════════
#  REPORT PREVIEW
# ═══════════════════════════════════════════════

# ── MASTHEAD ──────────────────────────────────
st.markdown(f"""
<div class="rpt-masthead">
  <div class="rpt-masthead-left">
    <div class="rt">Player Comparison Report</div>
    <div class="rs">{sel_pa}  &nbsp;vs&nbsp;
      {sel_pb}  &nbsp;·&nbsp;
      {sel_role} &nbsp;·&nbsp; {sel_yr}</div>
  </div>
  <div class="rpt-meta-box">
    <div class="rdate">Generated: {gen_ts}</div>
    <div class="rid">{rpt_id}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── SUMMARY BOX ───────────────────────────────
st.markdown(f"""
<div class="rpt-summary">
  <div class="sum-item">
    <div class="sl">Player A</div>
    <div class="sv">{sel_pa} <small style="font-size:.7rem;color:#4A6FA5;"></small></div>
  </div>
  <div class="sum-item">
    <div class="sl">Player B</div>
    <div class="sv">{sel_pb} <small style="font-size:.7rem;color:#4A6FA5;"></small></div>
  </div>
  <div class="sum-item">
    <div class="sl">Role Focus</div>
    <div class="sv">{sel_role}</div>
  </div>
  <div class="sum-item">
    <div class="sl">Period</div>
    <div class="sv">{sel_yr}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  PAGE 1 — PERFORMANCE COMPARISON
# ══════════════════════════════════════════════

sec_num = 1

# ── SECTION: BATTING (only if role warrants) ──
if show_bat:
    st.markdown(sec(sec_num,"","Batting Performance Comparison"),
                unsafe_allow_html=True)
    sec_num += 1
    rows_a = ""; rows_b = ""
    for lbl, v1, v2, hb in build_bat_metrics(bsa, bsb):
        r1, r2 = cmp_row(lbl, v1, v2, hb)
        rows_a += r1; rows_b += r2
    st.markdown(f"""
    <div class="cmp-grid">
      <div class="cmp-col p1-col">
        <div class="cmp-player-name">{sel_pa}</div>
        {rows_a}
      </div>
      <div class="cmp-col p2-col">
        <div class="cmp-player-name p2-name">{sel_pb}</div>
        {rows_b}
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown(insight_block(ins_bat), unsafe_allow_html=True)

# ── SECTION: BOWLING (only if role warrants) ──
if show_bowl:
    st.markdown('<hr class="hr">', unsafe_allow_html=True)
    st.markdown(sec(sec_num,"","Bowling Performance Comparison"),
                unsafe_allow_html=True)
    sec_num += 1
    rows_a = ""; rows_b = ""
    for lbl, v1, v2, hb in build_bowl_metrics(bwa, bwb):
        r1, r2 = cmp_row(lbl, v1, v2, hb)
        rows_a += r1; rows_b += r2
    st.markdown(f"""
    <div class="cmp-grid">
      <div class="cmp-col p1-col">
        <div class="cmp-player-name">{sel_pa}</div>
        {rows_a}
      </div>
      <div class="cmp-col p2-col">
        <div class="cmp-player-name p2-name">{sel_pb}</div>
        {rows_b}
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown(insight_block(ins_bowl), unsafe_allow_html=True)

# ── SECTION: RECENT FORM ──────────────────────
st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec(sec_num,"","Recent Form Comparison (Last 5 Matches)"),
            unsafe_allow_html=True)
sec_num += 1

rows_a = ""; rows_b = ""
for lbl, v1, v2, hb in build_form_metrics(fb_a_bat, fb_b_bat,
                                           fb_a_bowl, fb_b_bowl, sel_role):
    r1, r2 = cmp_row(lbl, v1, v2, hb)
    rows_a += r1; rows_b += r2

st.markdown(f"""
<div class="cmp-grid">
  <div class="cmp-col p1-col">
    <div class="cmp-player-name">{sel_pa}</div>
    {rows_a if rows_a else '<div class="cmp-stat"><span class="cmp-stat-lbl">No recent data</span></div>'}
  </div>
  <div class="cmp-col p2-col">
    <div class="cmp-player-name p2-name">{sel_pb}</div>
    {rows_b if rows_b else '<div class="cmp-stat"><span class="cmp-stat-lbl">No recent data</span></div>'}
  </div>
</div>
""", unsafe_allow_html=True)
st.markdown(single_insight(form_ins), unsafe_allow_html=True)

# ══════════════════════════════════════════════
#  PAGE 2 — ADVANCED COMPARISON
# ══════════════════════════════════════════════
# ── SECTION: OPPONENT ────────────────────────
opp_label_val = "Best Opp. Avg" if rl_norm_combined != "bowler" else "Wickets vs Best Opp."
opp_r1, opp_r2 = cmp_row(opp_label_val, opp_a_val, opp_b_val, True)
st.markdown(sec(sec_num,"","Opponent Performance Comparison"),
            unsafe_allow_html=True)
sec_num += 1
st.markdown(f"""
<div class="cmp-grid">
  <div class="cmp-col p1-col">
    <div class="cmp-player-name">{sel_pa}</div>
    <div class="cmp-stat"><span class="cmp-stat-lbl">Best Opponent</span>
      <span class="cmp-stat-val">{opp_a}</span></div>
    {opp_r1}
  </div>
  <div class="cmp-col p2-col">
    <div class="cmp-player-name p2-name">{sel_pb}</div>
    <div class="cmp-stat"><span class="cmp-stat-lbl">Best Opponent</span>
      <span class="cmp-stat-val">{opp_b}</span></div>
    {opp_r2}
  </div>
</div>
""", unsafe_allow_html=True)
st.markdown(single_insight(opp_ins), unsafe_allow_html=True)

# ── SECTION: VENUE ───────────────────────────
ven_label_val = "Best Venue Avg" if rl_norm_combined != "bowler" else "Wickets at Best Venue"
ven_r1, ven_r2 = cmp_row(ven_label_val, ven_a_val, ven_b_val, True)
st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec(sec_num,"","Venue Performance Comparison"),
            unsafe_allow_html=True)
sec_num += 1
st.markdown(f"""
<div class="cmp-grid">
  <div class="cmp-col p1-col">
    <div class="cmp-player-name">{sel_pa}</div>
    <div class="cmp-stat"><span class="cmp-stat-lbl">Best Venue</span>
      <span class="cmp-stat-val">{ven_a}</span></div>
    {ven_r1}
  </div>
  <div class="cmp-col p2-col">
    <div class="cmp-player-name p2-name">{sel_pb}</div>
    <div class="cmp-stat"><span class="cmp-stat-lbl">Best Venue</span>
      <span class="cmp-stat-val">{ven_b}</span></div>
    {ven_r2}
  </div>
</div>
""", unsafe_allow_html=True)
st.markdown(single_insight(ven_ins), unsafe_allow_html=True)

# ── SECTION: STRENGTHS ───────────────────────
st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec(sec_num,"","Strength Comparison"),
            unsafe_allow_html=True)
sec_num += 1
str_a_html = "".join(sw_item(s, "-") for s in str_a)
str_b_html = "".join(sw_item(s, "-") for s in str_b)
st.markdown(f"""
<div class="sw-grid">
  <div class="sw-box strength">
    <div class="sw-box-title">{sel_pa} — Strengths</div>
    {str_a_html}
  </div>
  <div class="sw-box strength">
    <div class="sw-box-title">{sel_pb} — Strengths</div>
    {str_b_html}
  </div>
</div>
""", unsafe_allow_html=True)

# ── SECTION: WEAKNESSES ──────────────────────
st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec(sec_num,"","Weakness Comparison"),
            unsafe_allow_html=True)
sec_num += 1
wk_a_html = "".join(sw_item(w, "!") for w in wk_a)
wk_b_html = "".join(sw_item(w, "!") for w in wk_b)
st.markdown(f"""
<div class="sw-grid">
  <div class="sw-box weakness">
    <div class="sw-box-title">{sel_pa} — Areas to Improve</div>
    {wk_a_html}
  </div>
  <div class="sw-box weakness">
    <div class="sw-box-title">{sel_pb} — Areas to Improve</div>
    {wk_b_html}
  </div>
</div>
""", unsafe_allow_html=True)

# ── SECTION: KEY INSIGHTS ────────────────────
st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec(sec_num,"","Key Comparison Insights"),
            unsafe_allow_html=True)
sec_num += 1
ki_html = "".join(key_insight_item(i+1, t) for i, t in enumerate(key_ins))
st.markdown(ki_html, unsafe_allow_html=True)

# ── SECTION: VERDICT ─────────────────────────
st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec(sec_num,"","Final Comparison Verdict"),
            unsafe_allow_html=True)
sec_num += 1
st.markdown(f"""
<div class="verdict-box">
  <div class="verdict-title">Analyst Verdict</div>
  <div class="verdict-text">{verdict}</div>
</div>
""", unsafe_allow_html=True)
# ═══════════════════════════════════════════════
#  PDF BUILDER  — clean minimal layout
# ═══════════════════════════════════════════════
def build_pdf() -> bytes:
    buf = io.BytesIO()

    C_NAVY  = colors.HexColor("#1B3A6B")
    C_AMBER = colors.HexColor("#E8A020")
    C_STEEL = colors.HexColor("#4A6FA5")
    C_GREEN = colors.HexColor("#1E7A34")
    C_RED   = colors.HexColor("#C0392B")
    C_LIGHT = colors.HexColor("#EBF3FF")
    C_BG    = colors.HexColor("#f7f9fc")
    C_BORD  = colors.HexColor("#dde3ee")
    C_WHITE = colors.white

    def on_page(canvas, doc):
        W, H = A4
        canvas.saveState()
        canvas.setStrokeColor(C_NAVY)
        canvas.setLineWidth(1.2)
        canvas.rect(1.2*cm, 1.2*cm, W - 2.4*cm, H - 2.4*cm)
        canvas.setFont("Times-Roman", 8)
        canvas.setFillColor(C_STEEL)
        canvas.drawRightString(W - 1.5*cm, 1.0*cm,
            f"Page {doc.page}  ·  {sel_pa} vs {sel_pb}  ·  {rpt_id}")
        canvas.restoreState()

    doc = SimpleDocTemplate(buf, pagesize=A4,
        leftMargin=2.0*cm, rightMargin=2.0*cm,
        topMargin=2.0*cm,  bottomMargin=2.2*cm,
        title=f"Comparison Report — {sel_pa} vs {sel_pb}")

    # ── Style helpers ───────────────────────────
    def PS(name, **kw):
        from reportlab.lib.styles import getSampleStyleSheet
        p = getSampleStyleSheet()["Normal"]
        return ParagraphStyle(name, parent=p, **kw)

    # Base styles
    S_NORMAL  = PS("N",  fontName="Times-Roman", fontSize=10, leading=15)
    S_BOLD    = PS("B",  fontName="Times-Bold",  fontSize=10, leading=15)
    S_SMALL   = PS("SM", fontName="Times-Roman", fontSize=9,  textColor=C_STEEL, leading=13)
    S_WHITE   = PS("WH", fontName="Times-Bold",  fontSize=15, textColor=C_WHITE, leading=22)
    S_META    = PS("ME", fontName="Times-Roman", fontSize=9,  textColor=colors.HexColor("#bbbbbb"),
                   alignment=TA_RIGHT, leading=13)
    S_SEC     = PS("SC", fontName="Times-Bold",  fontSize=11, textColor=C_NAVY,
                   spaceBefore=10, spaceAfter=3, leading=16)
    S_INS     = PS("IN", fontName="Times-Roman", fontSize=9,  textColor=colors.HexColor("#6B4B00"),
                   leftIndent=6, spaceAfter=2, leading=14)
    S_GREEN   = PS("GR", fontName="Times-Roman", fontSize=10, textColor=C_GREEN,
                   leftIndent=8, spaceAfter=2, leading=14)
    S_RED     = PS("RD", fontName="Times-Roman", fontSize=10, textColor=C_RED,
                   leftIndent=8, spaceAfter=2, leading=14)
    S_GBOLD   = PS("GB", fontName="Times-Bold",  fontSize=10, textColor=C_GREEN, leading=14)
    S_RBOLD   = PS("RB", fontName="Times-Bold",  fontSize=10, textColor=C_RED, leading=14)
    S_KI      = PS("KI", fontName="Times-Roman", fontSize=10, leftIndent=8,
                   spaceAfter=3, leading=14)
    S_VDT     = PS("VD", fontName="Times-Roman", fontSize=10, textColor=C_WHITE,
                   leading=16, alignment=TA_JUSTIFY)

    # ── Reusable element builders ────────────────

    def section_heading(num, title):
        """Plain bold section heading + navy underline — no number prefix."""
        return [
            Paragraph(title, S_SEC),
            HRFlowable(width="100%", thickness=0.8, color=C_NAVY, spaceAfter=4),
        ]

    def stat_table(metrics, p1_name, p2_name):
        """
        3-col table: Metric | P1 value | P2 value.
        Green = better, Red = worse (per higher_better flag).
        """
        hdr = [
            Paragraph("Metric", PS("TH", fontName="Times-Bold", fontSize=10,
                                   textColor=C_WHITE)),
            Paragraph(p1_name,  PS("TH1", fontName="Times-Bold", fontSize=10,
                                   textColor=C_WHITE, alignment=TA_CENTER)),
            Paragraph(p2_name,  PS("TH2", fontName="Times-Bold", fontSize=10,
                                   textColor=C_WHITE, alignment=TA_CENTER)),
        ]
        rows = [hdr]
        for lbl, v1, v2, hb in metrics:
            try:
                n1 = float(str(v1).replace("*", "").replace("%", ""))
                n2 = float(str(v2).replace("*", "").replace("%", ""))
                if n1 == n2:
                    c1 = c2 = C_NAVY
                elif (hb and n1 > n2) or (not hb and n1 < n2):
                    c1, c2 = C_GREEN, C_RED
                else:
                    c1, c2 = C_RED, C_GREEN
            except:
                c1 = c2 = colors.HexColor("#2c3e50")

            def cell(v, clr):
                return Paragraph(str(v),
                    PS(f"CV{id(v)}", fontName="Times-Bold", fontSize=10,
                       textColor=clr, alignment=TA_CENTER, leading=14))

            rows.append([Paragraph(lbl, S_SMALL), cell(v1, c1), cell(v2, c2)])

        t = Table(rows, colWidths=["46%", "27%", "27%"])
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), C_NAVY),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [C_WHITE, C_BG]),
            ("LINEBELOW",     (0, 0), (-1, -1), 0.3, C_BORD),
            ("BOX",           (0, 0), (-1, -1), 0.5, C_BORD),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ]))
        return t

    def two_col_text(left_items, right_items, left_title, right_title,
                     title_style, item_style):
        """Side-by-side two-column list (strengths or weaknesses)."""
        left  = [Paragraph(left_title,  title_style)]  + [Paragraph(s, item_style) for s in left_items]
        right = [Paragraph(right_title, title_style)]  + [Paragraph(s, item_style) for s in right_items]
        mx    = max(len(left), len(right))
        left  += [Spacer(1, 1)] * (mx - len(left))
        right += [Spacer(1, 1)] * (mx - len(right))
        t = Table(list(zip(left, right)), colWidths=["50%", "50%"])
        t.setStyle(TableStyle([
            ("VALIGN",       (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING",   (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
            ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ]))
        return t

    def insight_paras(items):
        return [Paragraph(f"•  {strip_html(i)}", S_INS) for i in items]

    def hr():
        return HRFlowable(width="100%", thickness=0.4,
                          color=colors.HexColor("#cccccc"), spaceAfter=6, spaceBefore=8)

    # ── BUILD STORY ─────────────────────────────
    story = []

    # ─ MASTHEAD (navy block — keep as requested) ─
    mhd = Table([[
        Paragraph("Player Comparison Report", S_WHITE),
        Paragraph(f"ID: {rpt_id}<br/>Generated: {gen_ts}", S_META),
    ]], colWidths=["65%", "35%"])
    mhd.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), C_NAVY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("LEFTPADDING",   (0, 0), (-1, -1), 14),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
    ]))
    story += [mhd, Spacer(1, 8)]

    # ─ SUMMARY LINE ─────────────────────────────
    story.append(Paragraph(
        f"<b>Player A:</b> {sel_pa}  &nbsp;&nbsp;  "
        f"<b>Player B:</b> {sel_pb}  &nbsp;&nbsp;  "
        f"<b>Role Focus:</b> {sel_role}  &nbsp;&nbsp;  "
        f"<b>Period:</b> {sel_yr}",
        PS("SUM", fontName="Times-Roman", fontSize=9, textColor=C_STEEL,
           leading=14, spaceAfter=10)))
    story.append(HRFlowable(width="100%", thickness=1, color=C_NAVY, spaceAfter=10))

    sec_num = 1

    # ─ BATTING ──────────────────────────────────
    if show_bat:
        story += section_heading(sec_num, "Batting Performance Comparison")
        sec_num += 1
        story.append(stat_table(build_bat_metrics(bsa, bsb), sel_pa, sel_pb))
        story.append(Spacer(1, 4))
        story += insight_paras(ins_bat)
        story.append(hr())

    # ─ BOWLING ──────────────────────────────────
    if show_bowl:
        story += section_heading(sec_num, "Bowling Performance Comparison")
        sec_num += 1
        story.append(stat_table(build_bowl_metrics(bwa, bwb), sel_pa, sel_pb))
        story.append(Spacer(1, 4))
        story += insight_paras(ins_bowl)
        story.append(hr())

    # ─ RECENT FORM ──────────────────────────────
    story += section_heading(sec_num, "Recent Form Comparison (Last 5 Matches)")
    sec_num += 1
    form_m = build_form_metrics(fb_a_bat, fb_b_bat, fb_a_bowl, fb_b_bowl, sel_role)
    story.append(stat_table(form_m, sel_pa, sel_pb))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"•  {strip_html(form_ins)}", S_INS))
    story.append(hr())

    # ─ OPPONENT ─────────────────────────────────
    opp_lbl = "Best Opp. Avg" if rl_norm_combined != "bowler" else "Wickets vs Best Opp."
    story += section_heading(sec_num, "Opponent Performance Comparison")
    sec_num += 1
    story.append(stat_table([
        ("Best Opponent", opp_a, opp_b, True),
        (opp_lbl,        opp_a_val, opp_b_val, True),
    ], sel_pa, sel_pb))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"•  {strip_html(opp_ins)}", S_INS))
    story.append(hr())

    # ─ VENUE ────────────────────────────────────
    ven_lbl = "Best Venue Avg" if rl_norm_combined != "bowler" else "Wickets at Best Venue"
    story += section_heading(sec_num, "Venue Performance Comparison")
    sec_num += 1
    story.append(stat_table([
        ("Best Venue", ven_a, ven_b, True),
        (ven_lbl,      ven_a_val, ven_b_val, True),
    ], sel_pa, sel_pb))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"•  {strip_html(ven_ins)}", S_INS))
    story.append(hr())

    # ─ STRENGTHS ────────────────────────────────
    story += section_heading(sec_num, "Strength Comparison")
    sec_num += 1
    story.append(two_col_text(
        left_items  =[f"-  {s}" for s in str_a],
        right_items =[f"-  {s}" for s in str_b],
        left_title  =f"{sel_pa}",
        right_title =f"{sel_pb}",
        title_style =PS("STH", fontName="Times-Bold", fontSize=10,
                        textColor=C_GREEN, leading=14, spaceAfter=3),
        item_style  =S_GREEN,
    ))
    story.append(hr())

    # ─ WEAKNESSES ───────────────────────────────
    story += section_heading(sec_num, "Weakness Comparison")
    sec_num += 1
    story.append(two_col_text(
        left_items  =[f"!  {w}" for w in wk_a],
        right_items =[f"!  {w}" for w in wk_b],
        left_title  =f"{sel_pa}",
        right_title =f"{sel_pb}",
        title_style =PS("WKH", fontName="Times-Bold", fontSize=10,
                        textColor=C_RED, leading=14, spaceAfter=3),
        item_style  =S_RED,
    ))
    story.append(hr())

    # ─ KEY INSIGHTS ─────────────────────────────
    story += section_heading(sec_num, "Key Comparison Insights")
    sec_num += 1
    for i, ki in enumerate(key_ins, 1):
        story.append(Paragraph(f"{i}.  {strip_html(ki)}", S_KI))
    story.append(hr())

    # ─ VERDICT ──────────────────────────────────
    story += section_heading(sec_num, "Final Comparison Verdict")
    sec_num += 1
    story.append(Paragraph(
        f"<b>Analyst Verdict:</b>  {strip_html(verdict)}",
        PS("VDT", fontName="Times-Roman", fontSize=10, textColor=colors.HexColor("#1a1a1a"),
           leading=17, spaceAfter=6, alignment=TA_JUSTIFY)))
    story.append(Spacer(1, 14))

    # ─ FOOTER ───────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.8, color=C_NAVY, spaceAfter=4))

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════
#  DOWNLOAD BAR
# ═══════════════════════════════════════════════
st.markdown('<div class="dl-bar">', unsafe_allow_html=True)
dl_a, dl_b, dl_c = st.columns([3.5, 3, 1.5])

with dl_a:
    st.markdown(
        f'<div class="dl-info"><strong>{sel_pa}</strong> vs '
        f'<strong>{sel_pb}</strong><br>'
        f'{sel_role} &nbsp;·&nbsp; {sel_yr} &nbsp;·&nbsp; {gen_ts}</div>',
        unsafe_allow_html=True)

with dl_b:
    st.markdown("", unsafe_allow_html=True)

with dl_c:
    pdf_bytes = build_pdf()
    fname     = (f"CMP_{sel_pa.replace(' ','_')}_vs_"
                 f"{sel_pb.replace(' ','_')}_{sel_yr}.pdf")
    st.download_button(
        label="Download PDF",
        data=pdf_bytes, file_name=fname,
        mime="application/pdf",
        use_container_width=True)

# ── SITE FOOTER ───────────────────────────────
st.markdown("""
<div class="site-footer">
  Cricket Player Performance Analysis Dashboard &nbsp;|&nbsp; Built with Streamlit
</div>
""", unsafe_allow_html=True)