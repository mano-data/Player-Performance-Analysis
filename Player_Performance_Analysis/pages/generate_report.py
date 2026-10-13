import streamlit as st
import pandas as pd
import openpyxl, base64, os, io, re
from datetime import datetime

# ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, HRFlowable, KeepTogether
)

# ═══════════════════════════════════════════════════════════════
#  PAGE CONFIG
# ═══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Report Generator | Cricket Analytics",
    page_icon="", layout="wide",
    initial_sidebar_state="collapsed",
)

# ═══════════════════════════════════════════════════════════════
#  PATHS
# ═══════════════════════════════════════════════════════════════
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR    = os.path.dirname(BASE_DIR)
DATASET_DIR = os.path.join(ROOT_DIR, "Dataset")
IMAGE_DIR   = os.path.join(ROOT_DIR, "images")

def _find(fname):
    for d in [DATASET_DIR, ROOT_DIR, BASE_DIR]:
        p = os.path.join(d, fname)
        if os.path.exists(p):
            return p
    return os.path.join(DATASET_DIR, fname)

MASTER_PATH = _find("player_master.csv")
STATS_PATH  = _find("player_stats2.xlsx")

# ═══════════════════════════════════════════════════════════════
#  BRAND PALETTE
# ═══════════════════════════════════════════════════════════════
NAVY   = "#1B3A6B"
AMBER  = "#E8A020"
STEEL  = "#4A6FA5"
GREEN  = "#1E7A34"
RED    = "#C0392B"
LIGHT  = "#EBF3FF"
BGRAY  = "#f7f9fc"
BORDER = "#dde3ee"
MUTED  = "#6B7A8D"

SENA_SET = {"South Africa", "England", "New Zealand", "Australia"}
ASIA_SET = {"India", "Sri Lanka", "Bangladesh", "Pakistan", "UAE"}

ZONE_MAP = {
    "Long On":    "Runs_Longon",
    "Long Off":   "Runs_Longoff",
    "Cover":      "Runs_Covers",
    "Point":      "Runs_Point",
    "Third Man":  "Runs_Thirdman",
    "Fine Leg":   "Runs_Fineleg",
    "Square Leg": "Runs_Square",
    "Mid Wicket": "Runs_Midwct",
}

# ═══════════════════════════════════════════════════════════════
#  GLOBAL CSS
# ═══════════════════════════════════════════════════════════════
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

.navbar{
    background:#fff;border-bottom:2px solid #1B3A6B;padding:0 60px;height:64px;
    display:flex;align-items:center;justify-content:space-between;
    position:fixed;width:100%;top:0;z-index:999;
    box-shadow:0 2px 14px rgba(27,58,107,.09);margin:0 -60px 0 -60px;
}
.nav-brand{font-family:'Poppins',sans-serif;font-size:1.18rem;font-weight:700;color:#1B3A6B;
    display:flex;align-items:center;gap:8px;}
.nav-brand span{color:#E8A020;}
.nav-right{display:flex;gap:8px;align-items:center;}
.nb{display:inline-flex;align-items:center;gap:5px;padding:7px 16px;border-radius:6px;
    font-size:.82rem;font-weight:600;text-decoration:none !important;font-family:'Inter',sans-serif;
    cursor:pointer;transition:all .18s;}
.nb-ghost{color:#1B3A6B !important;border:1.5px solid #1B3A6B;background:transparent;}
.nb-ghost:hover{background:#EBF3FF;}
.nb-solid{background:#1B3A6B;color:#fff !important;border:1.5px solid #1B3A6B;}
.nb-solid:hover{background:#243f7a;}

.ptb{background:#f7f9fc;border-left:5px solid #1B3A6B;padding:14px 60px;margin:40px -60px 28px -60px;}
.ptb h2{font-family:'Poppins',sans-serif;font-size:1.15rem;font-weight:700;color:#1B3A6B;margin:0;}
.ptb p{color:#6B7A8D;font-size:.8rem;margin:3px 0 0 0;}

.filter-panel{background:#f7f9fc;border:1px solid #dde3ee;border-radius:12px;
    padding:22px 28px 18px;margin-bottom:28px;}
.fl{font-size:.74rem;font-weight:700;color:#1B3A6B;letter-spacing:.8px;
    text-transform:uppercase;margin-bottom:5px;}
.stSelectbox>div>div{background:#fff !important;border:1.5px solid #c8d0e0 !important;
    border-radius:7px !important;color:#1B3A6B !important;font-size:.88rem !important;}
div[data-testid="stButton"]>button{background:#1B3A6B !important;color:#fff !important;
    font-family:'Poppins',sans-serif !important;font-weight:700 !important;
    font-size:.9rem !important;letter-spacing:.8px !important;border:none !important;
    border-radius:7px !important;padding:12px 0 !important;width:100% !important;}
div[data-testid="stButton"]>button:hover{background:#243f7a !important;}
div[data-testid="stDownloadButton"]>button{background:#1E7A34 !important;color:#fff !important;
    font-family:'Poppins',sans-serif !important;font-weight:700 !important;
    font-size:.88rem !important;border:none !important;border-radius:7px !important;
    padding:11px 0 !important;width:100% !important;}
div[data-testid="stDownloadButton"]>button:hover{background:#165e28 !important;}

.rpt-wrap{background:#fff;border:2px solid #1B3A6B;border-radius:14px;padding:48px 56px;
    box-shadow:0 8px 48px rgba(27,58,107,.12);max-width:980px;margin:0 auto 32px auto;}

.masthead{background:linear-gradient(135deg,#1B3A6B 0%,#22488a 100%);border-radius:10px;
    padding:24px 30px;margin-bottom:32px;display:flex;align-items:center;
    justify-content:space-between;}
.masthead-left{display:flex;align-items:center;gap:14px;}
.masthead-icon{width:46px;height:46px;border-radius:50%;background:#E8A020;
    display:flex;align-items:center;justify-content:center;font-size:1.3rem;flex-shrink:0;}
.masthead-text .mt{font-family:'Poppins',sans-serif;font-size:1.15rem;font-weight:800;color:#fff;}
.masthead-text .ms{font-size:.78rem;color:rgba(255,255,255,.6);margin-top:2px;}
.masthead-right{text-align:right;}
.masthead-right .mr-date{font-size:.75rem;color:rgba(255,255,255,.5);}
.masthead-right .mr-id{display:inline-block;margin-top:6px;background:rgba(255,255,255,.12);
    border-radius:20px;padding:2px 12px;font-size:.7rem;font-weight:600;
    color:rgba(255,255,255,.75);letter-spacing:.4px;}

.rpt-sec-head{display:flex;align-items:center;gap:10px;background:#EBF3FF;
    border-left:5px solid #1B3A6B;border-radius:0 8px 8px 0;
    padding:10px 16px;margin:28px 0 16px 0;}
.sec-num{width:26px;height:26px;border-radius:50%;background:#1B3A6B;color:#fff;
    display:flex;align-items:center;justify-content:center;
    font-family:'Poppins',sans-serif;font-size:.75rem;font-weight:800;flex-shrink:0;}
.sec-label{font-family:'Poppins',sans-serif;font-size:1rem;font-weight:700;color:#1B3A6B;}
.rpt-subsec{font-family:'Poppins',sans-serif;font-size:.82rem;font-weight:700;color:#4A6FA5;
    text-transform:uppercase;letter-spacing:.7px;padding:7px 14px;
    border-left:3px solid #E8A020;background:#fffaf0;border-radius:0 6px 6px 0;
    margin:14px 0 10px 0;}

.profile-outer{background:#f7f9fc;border:1px solid #dde3ee;border-radius:12px;
    padding:24px;display:flex;gap:30px;align-items:flex-start;}
.profile-img-col{flex:0 0 130px;text-align:center;}
.profile-img-col img{width:110px;height:110px;border-radius:50%;object-fit:cover;
    border:3px solid #1B3A6B;display:block;margin:0 auto;}
.player-avatar{width:110px;height:110px;border-radius:50%;background:#EBF3FF;
    border:3px solid #1B3A6B;display:flex;align-items:center;justify-content:center;
    margin:0 auto;font-family:'Poppins',sans-serif;font-size:2.2rem;
    font-weight:800;color:#1B3A6B;}
.pname{font-family:'Poppins',sans-serif;font-size:1rem;font-weight:700;
    color:#1B3A6B;margin-top:10px;text-align:center;}
.role-badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:.68rem;
    font-weight:700;letter-spacing:.4px;text-transform:uppercase;margin-top:4px;}
.rb-bat{background:#EBF3FF;color:#1B3A6B;border:1px solid #b0c8e8;}
.rb-bowl{background:#FFF0E8;color:#C0501A;border:1px solid #f0c0a0;}
.rb-ar{background:#F0F8EE;color:#2A7A2A;border:1px solid #a8d4a0;}

.profile-info-col{flex:1;}
.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:2px 24px;}
.info-row{display:flex;align-items:center;gap:8px;padding:7px 0;
    border-bottom:1px solid #f0f3f8;font-size:.82rem;}
.info-row:last-child{border-bottom:none;}
.il{color:#8899aa;font-weight:500;min-width:108px;}
.iv{color:#1B3A6B;font-weight:600;}

.filter-chips-box{background:#fff8ee;border:1px solid #f0d090;border-radius:8px;
    padding:11px 16px;margin-top:16px;}
.fct{font-size:.7rem;font-weight:700;color:#8A6000;text-transform:uppercase;
    letter-spacing:.6px;margin-bottom:7px;}
.chip{display:inline-block;background:#fff;border:1px solid #dde3ee;border-radius:20px;
    padding:3px 11px;font-size:.72rem;font-weight:600;color:#4A6FA5;margin:2px 3px;}

.kpi-band{display:flex;flex-wrap:wrap;gap:10px;margin:12px 0;}
.kpi-c{flex:1;min-width:95px;background:#fff;border:1px solid #dde3ee;border-radius:10px;
    padding:14px 10px;text-align:center;box-shadow:0 2px 8px rgba(27,58,107,.05);
    position:relative;overflow:hidden;}
.kpi-c::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;background:#1B3A6B;}
.kpi-c.c-amber::after{background:#E8A020;}
.kpi-c.c-green::after{background:#1E7A34;}
.kpi-v{font-family:'Poppins',sans-serif;font-size:1.45rem;font-weight:700;
    color:#1B3A6B;line-height:1;}
.kpi-l{font-size:.65rem;color:#8899aa;font-weight:600;text-transform:uppercase;
    letter-spacing:.5px;margin-top:6px;}

.ins{display:flex;align-items:flex-start;gap:11px;padding:10px 14px;border-radius:8px;
    margin-bottom:7px;font-size:.84rem;line-height:1.6;}
.ins-pos{background:#f0f9f2;border-left:4px solid #1E7A34;}
.ins-neg{background:#fdf0f0;border-left:4px solid #C0392B;}
.ins-neu{background:#f7f9fc;border-left:4px solid #E8A020;}
.ins-icon{font-size:1rem;flex-shrink:0;margin-top:2px;}
.ins-body{color:#2c3e50;}
.ins-body strong{color:#1B3A6B;}

.ft{width:100%;border-collapse:collapse;font-size:.8rem;margin-top:10px;}
.ft thead th{background:#1B3A6B;color:#fff;font-family:'Poppins',sans-serif;
    font-size:.7rem;font-weight:700;letter-spacing:.5px;padding:9px 12px;text-align:center;}
.ft tbody td{padding:8px 12px;border-bottom:1px solid #f0f3f8;text-align:center;color:#2c3e50;}
.ft tbody tr:last-child td{border-bottom:none;}
.ft tbody tr:hover td{background:#fafbfd;}
.res-w{color:#1E7A34;font-weight:700;}
.res-l{color:#C0392B;font-weight:700;}
.res-n{color:#8899aa;font-weight:600;}

.rpt-foot{background:#1B3A6B;border-radius:8px;padding:16px 24px;margin-top:34px;
    display:flex;align-items:center;justify-content:space-between;}
.rf-brand{font-family:'Poppins',sans-serif;font-size:.88rem;font-weight:700;color:#fff;}
.rf-brand span{color:#E8A020;}
.rf-meta{font-size:.73rem;color:rgba(255,255,255,.5);text-align:right;}

.dl-bar{background:#f7f9fc;border-top:1px solid #dde3ee;padding:22px 60px;
    margin:0 -60px 0 -60px;display:flex;align-items:center;
    justify-content:space-between;flex-wrap:wrap;gap:14px;}
.dl-info{font-size:.82rem;color:#6B7A8D;line-height:1.7;}
.dl-info strong{color:#1B3A6B;}

.hr{border:none;height:1px;background:#e8edf5;margin:22px 0;}

.nodata{text-align:center;padding:70px 20px;color:#8899aa;}
.nodata .nd-icon{font-size:3.2rem;display:block;margin-bottom:12px;}
.nodata p{font-size:.9rem;margin:0;}

.pg-foot{background:#1B3A6B;color:rgba(255,255,255,.7);text-align:center;
    padding:18px 60px;font-size:.8rem;margin:0 -60px -60px -60px;letter-spacing:.3px;}
.pg-foot span{color:#E8A020;font-weight:600;}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
#  DATA LOADING
# ═══════════════════════════════════════════════════════════════
NON_NUMERIC = {
    "Match_id","Match_date","Opponent","Venue","Country","Match_result",
    "Dismissed_method","Dismissed_Length","Dismissed_against",
    "Dismissed_by","Dismissed _by",
}

@st.cache_data
def load_data():
    master = pd.read_csv(MASTER_PATH)
    master.columns = master.columns.str.strip()
    wb    = openpyxl.load_workbook(STATS_PATH)
    stats = {}
    for sheet in wb.sheetnames:
        ws   = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            stats[sheet.strip()] = pd.DataFrame()
            continue
        cols = [str(c).strip() if c else f"_c{i}"
                for i, c in enumerate(rows[0])]
        df   = pd.DataFrame(rows[1:], columns=cols)
        df["Match_date"] = pd.to_datetime(df["Match_date"], errors="coerce")
        for c in df.columns:
            if c not in NON_NUMERIC:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        stats[sheet.strip()] = df
    return master, stats

try:
    master_df, stats_dict = load_data()
except Exception as e:
    st.error(f"Data load error: {e}")
    st.stop()

# ═══════════════════════════════════════════════════════════════
#  UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════
def get_img_b64(filename: str):
    for d in [IMAGE_DIR, os.path.join(ROOT_DIR, "images"), BASE_DIR]:
        p = os.path.join(d, filename)
        if os.path.exists(p):
            ext  = filename.rsplit(".", 1)[-1].lower()
            mime = "jpeg" if ext in ("jpg", "jpeg") else ext
            with open(p, "rb") as f:
                return f"data:image/{mime};base64,{base64.b64encode(f.read()).decode()}"
    return None

def avatar_html(name: str, size: int = 110) -> str:
    ini = "".join(w[0].upper() for w in name.split()[:2])
    return (f'<div class="player-avatar" '
            f'style="width:{size}px;height:{size}px;'
            f'font-size:{size//2.6:.0f}px;">{ini}</div>')

def calc_avg(runs_s, not_out_s, innings: int) -> float:
    tot = float(runs_s.sum())
    dis = max(int(innings) - int(not_out_s.sum()), 0)
    return round(tot / dis, 2) if dis > 0 else round(tot, 2)

def safe_pct(a, b) -> float:
    return round(a / max(b, 1) * 100, 1)

def win_pct(df) -> float:
    if df.empty or "Match_result" not in df.columns:
        return 0.0
    return round((df["Match_result"] == "Won").sum() / max(len(df), 1) * 100, 1)

def gcol(df, *names):
    for n in names:
        if n in df.columns:
            return df[n]
    return pd.Series(0, index=df.index)

def bat_rows(df):
    return df[df["Batted_Flag"] == 1] if "Batted_Flag" in df.columns else pd.DataFrame()

def bowl_rows(df):
    return df[df["Bowl_Flag"] == 1] if "Bowl_Flag" in df.columns else pd.DataFrame()

def strip_tags(txt: str) -> str:
    return re.sub(r"<[^>]+>", "", str(txt)).replace("&amp;", "&").replace("&nbsp;", " ")

def role_norm(r: str) -> str:
    r = r.lower()
    if "all" in r: return "ar"
    if "bowl" in r: return "bowl"
    return "bat"

def role_badge_cls(role: str) -> str:
    r = role.lower()
    if "all" in r: return "rb-ar"
    if "bowl" in r: return "rb-bowl"
    return "rb-bat"

# ═══════════════════════════════════════════════════════════════
#  STAT CALCULATORS
# ═══════════════════════════════════════════════════════════════
def bat_stats(bdf: pd.DataFrame) -> dict:
    if bdf.empty:
        return {}
    no_s  = bdf.get("Not_out", pd.Series(0, index=bdf.index))
    inn   = len(bdf)
    runs  = int(bdf["Runs"].sum())
    bf    = int(bdf["Balls_faced"].sum()) if "Balls_faced" in bdf.columns else 0
    avg   = calc_avg(bdf["Runs"], no_s, inn)
    sr    = round(bdf["Strike_Rate"].replace(0, float("nan")).mean(), 1)
    f4    = int(bdf["4s"].sum())
    f6    = int(bdf["6s"].sum())
    fif   = int(bdf["50s"].sum()) if "50s" in bdf.columns else 0
    hun   = int(bdf["100s"].sum()) if "100s" in bdf.columns else 0
    no    = int(no_s.sum())
    dots  = int(bdf["Dots"].sum()) if "Dots" in bdf.columns else 0
    sg    = int(bdf["1s_2s_3s"].sum()) if "1s_2s_3s" in bdf.columns else 0
    best_idx = bdf["Runs"].idxmax()
    best  = f"{int(bdf.loc[best_idx,'Runs'])}{'*' if int(bdf.loc[best_idx,'Not_out']) else ''}"
    bnd   = f4 * 4 + f6 * 6
    wp    = win_pct(bdf)
    return dict(inn=inn, runs=runs, bf=bf, avg=avg, sr=sr, f4=f4, f6=f6,
                fif=fif, hun=hun, no=no, dots=dots, sg=sg, best=best,
                bnd=bnd, wp=wp)

def bowl_stats(bdf: pd.DataFrame) -> dict:
    if bdf.empty:
        return {}
    inn  = len(bdf)
    tbl  = int(bdf["Balls_Bowled"].sum()) if "Balls_Bowled" in bdf.columns else 0
    twk  = int(bdf["Wickets"].sum())
    rc   = int(bdf["Runs_Conceded"].sum()) if "Runs_Conceded" in bdf.columns else 0
    mai  = int(bdf["Maidens"].sum()) if "Maidens" in bdf.columns else 0
    eco  = round(bdf["Economy"].replace(0, float("nan")).mean(), 2)
    dots = int(gcol(bdf, "Dot_Balls").sum())
    ext  = int(gcol(bdf, "Extras").sum())
    bsr  = round(tbl / max(twk, 1), 1)
    dp   = safe_pct(dots, tbl)
    rhbw = int(bdf["Wicket_RHB"].sum()) if "Wicket_RHB" in bdf.columns else 0
    lhbw = int(bdf["Wicket_LHB"].sum()) if "Wicket_LHB" in bdf.columns else 0
    fd   = int(bdf["Full_Dismissals"].sum()) if "Full_Dismissals" in bdf.columns else 0
    gd   = int(gcol(bdf, "Good_Dissmisals", "Good_Dismissals").sum())
    sd   = int(gcol(bdf, "Short_Dissmisals", "Short_Dismissals").sum())
    wp   = win_pct(bdf)
    if twk > 0:
        r = bdf.sort_values(["Wickets", "Runs_Conceded"],
                            ascending=[False, True]).iloc[0]
        bfig = f"{int(r['Wickets'])}/{int(r.get('Runs_Conceded', 0))}"
    else:
        bfig = "0/0"
    return dict(inn=inn, tbl=tbl, ovs=f"{tbl//6}.{tbl%6}", twk=twk,
                rc=rc, mai=mai, eco=eco, dots=dots, ext=ext, bsr=bsr,
                dp=dp, rhbw=rhbw, lhbw=lhbw, fd=fd, gd=gd, sd=sd,
                wp=wp, bfig=bfig)

# ═══════════════════════════════════════════════════════════════
#  INSIGHT GENERATORS
# ═══════════════════════════════════════════════════════════════
def batsman_insights(bdf: pd.DataFrame, raw_df: pd.DataFrame) -> list:
    if bdf.empty:
        return []
    ins   = []
    bs    = bat_stats(bdf)
    runs  = bs["runs"]; bf = bs["bf"]; avg = bs["avg"]
    f4    = bs["f4"];   f6 = bs["f6"]; fif = bs["fif"]; hun = bs["hun"]
    bnd   = bs["bnd"];  sg = bs["sg"]; dots = bs["dots"]

    if "Opponent" in bdf.columns:
        og = bdf.groupby("Opponent").apply(
            lambda g: calc_avg(g["Runs"],
                               g.get("Not_out", pd.Series(0, index=g.index)),
                               len(g)))
        if not og.empty:
            best_o = og.idxmax(); best_v = og.max()
            worst_o= og.idxmin(); worst_v= og.min()
            ins.append(("",
                f"Best batting average against <strong>{best_o}</strong> "
                f"(Average: <strong>{best_v}</strong>)", "pos"))
            if worst_o != best_o:
                ins.append(("",
                    f"Relatively lower average against "
                    f"<strong>{worst_o}</strong> "
                    f"(Average: <strong>{worst_v}</strong>) — "
                    f"area to improve", "neg"))

    if "Opponent" in bdf.columns:
        rg = bdf.groupby("Opponent")["Runs"].sum()
        if not rg.empty:
            mr_o = rg.idxmax()
            ins.append(("",
                f"Most runs scored against <strong>{mr_o}</strong> "
                f"({int(rg.max())} runs across "
                f"{int((bdf['Opponent']==mr_o).sum())} innings)", "neu"))

    bnd_pct = safe_pct(bnd, max(runs, 1))
    if bnd_pct >= 50:
        ins.append(("",
            f"Relies heavily on boundary scoring — "
            f"<strong>{bnd_pct}%</strong> of total runs "
            f"come from boundaries "
            f"({bnd} boundary runs of {runs} total)", "neg"))
    else:
        ins.append(("",
            f"Balanced scoring pattern — "
            f"<strong>{bnd_pct}%</strong> of runs from boundaries, "
            f"strong rotation through the middle", "pos"))

    zv = {k: int(bdf[v].sum()) for k, v in ZONE_MAP.items()
          if v in bdf.columns}
    if zv:
        best_z  = max(zv, key=zv.get)
        worst_z = min(zv, key=zv.get)
        ins.append(("",
            f"Scores most runs in the <strong>{best_z}</strong> zone "
            f"({zv[best_z]} runs) — primary hitting corridor", "pos"))
        if zv[worst_z] < zv[best_z] * 0.3:
            ins.append(("",
                f"Minimal scoring in the <strong>{worst_z}</strong> zone "
                f"({zv[worst_z]} runs) — a gap in range", "neu"))

    if "Country" in bdf.columns:
        home = bdf[bdf["Country"] == "India"]
        away = bdf[bdf["Country"] != "India"]
        if not home.empty and not away.empty:
            ha = calc_avg(home["Runs"],
                         home.get("Not_out", pd.Series(0, index=home.index)),
                         len(home))
            aa = calc_avg(away["Runs"],
                         away.get("Not_out", pd.Series(0, index=away.index)),
                         len(away))
            typ_ = "pos" if ha >= aa else "neg"
            ins.append(("",
                f"<strong>{'Performs better at home' if ha >= aa else 'Performs better away'}</strong> — "
                f"Home avg: <strong>{ha}</strong> | "
                f"Away avg: <strong>{aa}</strong>", typ_))

    if "Country" in bdf.columns:
        sena_df = bdf[bdf["Country"].isin(SENA_SET)]
        asia_df = bdf[bdf["Country"].isin(ASIA_SET)]
        if not sena_df.empty:
            sa = calc_avg(sena_df["Runs"],
                         sena_df.get("Not_out", pd.Series(0, index=sena_df.index)),
                         len(sena_df))
            ins.append(("",
                f"Average in SENA countries (SA/ENG/NZ/AUS): "
                f"<strong>{sa}</strong>", "neu"))
        if not asia_df.empty:
            aa2 = calc_avg(asia_df["Runs"],
                          asia_df.get("Not_out", pd.Series(0, index=asia_df.index)),
                          len(asia_df))
            ins.append(("",
                f"Average in Asian countries (IND/SL/BAN/PAK/UAE): "
                f"<strong>{aa2}</strong>", "neu"))

    if "Country" in bdf.columns:
        cg = bdf.groupby("Country").apply(
            lambda g: calc_avg(g["Runs"],
                               g.get("Not_out", pd.Series(0, index=g.index)),
                               len(g)))
        if not cg.empty:
            bc = cg.idxmax()
            ins.append(("",
                f"Performs best in <strong>{bc}</strong> with an average of "
                f"<strong>{cg.max()}</strong>", "pos"))

    if "Match_date" in bdf.columns:
        yg = bdf.groupby(bdf["Match_date"].dt.year).apply(
            lambda g: calc_avg(g["Runs"],
                               g.get("Not_out", pd.Series(0, index=g.index)),
                               len(g)))
        if len(yg) >= 2:
            trend = "improving" if yg.iloc[-1] >= yg.iloc[0] else "declining"
            typ_  = "pos" if trend == "improving" else "neg"
            ins.append(("",
                f"Year-on-year performance is <strong>{trend}</strong> "
                f"(from <strong>{yg.iloc[0]}</strong> in {int(yg.index[0])} "
                f"to <strong>{yg.iloc[-1]}</strong> in {int(yg.index[-1])})",
                typ_))

    if "Dismissed_against" in bdf.columns:
        dv = bdf[~bdf["Dismissed_against"].astype(str)
                   .isin(["Not Applicable", "0", "nan", ""])]
        if not dv.empty:
            top_d  = dv["Dismissed_against"].value_counts().index[0]
            top_cnt= dv["Dismissed_against"].value_counts().iloc[0]
            ins.append(("",
                f"Most dismissed by <strong>{top_d}</strong> "
                f"({top_cnt} times) — key vulnerability to address", "neg"))

    if "Dismissed_Length" in bdf.columns:
        dl = bdf[~bdf["Dismissed_Length"].astype(str)
                   .isin(["Not Applicable", "0", "nan", ""])]
        if not dl.empty:
            worst_len = dl["Dismissed_Length"].value_counts().index[0]
            cnt_len   = dl["Dismissed_Length"].value_counts().iloc[0]
            ins.append(("",
                f"Vulnerable to <strong>{worst_len}</strong> deliveries "
                f"({cnt_len} dismissals) — "
                f"opposition bowlers should note this", "neg"))

    if "Innings_no" in bdf.columns:
        i1 = bdf[bdf["Innings_no"] == 1]
        i2 = bdf[bdf["Innings_no"] == 2]
        if not i1.empty and not i2.empty:
            a1 = calc_avg(i1["Runs"],
                         i1.get("Not_out", pd.Series(0, index=i1.index)), len(i1))
            a2 = calc_avg(i2["Runs"],
                         i2.get("Not_out", pd.Series(0, index=i2.index)), len(i2))
            better = "1st" if a1 >= a2 else "2nd"
            typ_   = "pos" if a1 >= a2 else "neu"
            ins.append(("",
                f"Performs better in <strong>{better} innings</strong> — "
                f"1st innings avg: <strong>{a1}</strong> | "
                f"2nd innings avg: <strong>{a2}</strong>", typ_))

    return ins


def bowler_insights(bdf: pd.DataFrame) -> list:
    if bdf.empty:
        return []
    ins = []
    bs  = bowl_stats(bdf)

    ins.append(("",
        f"Bowling Pressure Index: <strong>{bs['dp']}%</strong> dot balls "
        f"({bs['dots']} of {bs['tbl']} balls) — "
        f"{'Elite level' if bs['dp'] >= 40 else 'Building'}",
        "pos" if bs["dp"] >= 40 else "neu"))

    bv = max(bs["fd"], bs["gd"], bs["sd"])
    bn = {bs["fd"]: "Full Length",
          bs["gd"]: "Good Length",
          bs["sd"]: "Short Pitch"}.get(bv, "Good Length")
    ins.append(("",
        f"Most effective with <strong>{bn}</strong> deliveries "
        f"({bv} dismissals) — primary weapon", "pos"))

    dom = "Right-Handed" if bs["rhbw"] >= bs["lhbw"] else "Left-Handed"
    ins.append(("",
        f"More effective vs <strong>{dom} batters</strong> "
        f"(RHB: {bs['rhbw']} wkts | LHB: {bs['lhbw']} wkts)", "pos"))

    if "Opponent" in bdf.columns:
        ow = bdf.groupby("Opponent")["Wickets"].sum()
        if not ow.empty:
            bo = ow.idxmax()
            ins.append(("",
                f"Most wickets against <strong>{bo}</strong> "
                f"({int(ow.max())} wickets)", "pos"))

    if "Opponent" in bdf.columns:
        oe = bdf.groupby("Opponent")["Economy"].mean().round(2)
        if not oe.empty:
            be_o = oe.idxmin()
            ins.append(("",
                f"Best economy against <strong>{be_o}</strong> "
                f"({oe.min()} RPO)", "pos"))

    if "Match_date" in bdf.columns:
        ye = bdf.groupby(bdf["Match_date"].dt.year)["Economy"].mean().round(2)
        if len(ye) >= 2:
            trend = "improving" if ye.iloc[-1] <= ye.iloc[0] else "declining"
            ins.append(("",
                f"Economy rate is <strong>{trend}</strong> "
                f"({ye.iloc[0]} in {int(ye.index[0])} to "
                f"{ye.iloc[-1]} in {int(ye.index[-1])})",
                "pos" if trend == "improving" else "neg"))

    if "Country" in bdf.columns:
        hb = bdf[bdf["Country"] == "India"]
        ab = bdf[bdf["Country"] != "India"]
        if not hb.empty and not ab.empty:
            he = round(hb["Economy"].mean(), 2)
            ae = round(ab["Economy"].mean(), 2)
            ins.append(("",
                f"Home economy: <strong>{he}</strong> | "
                f"Away economy: <strong>{ae}</strong> — "
                f"performs better <strong>"
                f"{'at home' if he <= ae else 'away'}</strong>",
                "pos" if he <= ae else "neg"))

    ins.append(("",
        f"Extras conceded: <strong>{bs['ext']}</strong> — "
        f"{'concern area' if bs['ext'] > 20 else 'well controlled'}",
        "neg" if bs["ext"] > 20 else "pos"))

    return ins


def key_insights_bat(bdf: pd.DataFrame) -> list:
    if bdf.empty:
        return []
    ins = []
    bs  = bat_stats(bdf)
    runs = bs["runs"]; bf = bs["bf"]
    fif  = bs["fif"]; hun = bs["hun"]
    dots = bs["dots"]; sg = bs["sg"]; bnd = bs["bnd"]

    conv = round(hun / max(fif + hun, 1) * 100, 1) if fif + hun > 0 else 0
    ins.append(("",
        f"50 to 100 Conversion rate: <strong>{conv}%</strong> "
        f"({hun} hundreds from {fif + hun} half-centuries+)",
        "pos" if conv >= 40 else "neu"))

    sr_pct = safe_pct(sg, max(runs, 1))
    ins.append(("",
        f"Strike rotation ability: <strong>{sr_pct}%</strong> of runs "
        f"via singles/twos/threes "
        f"({'excellent rotation' if sr_pct >= 30 else 'rotation can improve'})",
        "pos" if sr_pct >= 30 else "neg"))

    dp = safe_pct(dots, max(bf, 1))
    ins.append(("",
        f"Dot ball percentage: <strong>{dp}%</strong> of balls faced are dots "
        f"({'high pressure' if dp > 45 else 'well managed'})",
        "neg" if dp > 45 else "pos"))

    if "Match_result" in bdf.columns:
        bdf50 = bdf[bdf["Runs"] >= 50]
        if not bdf50.empty:
            wp50 = win_pct(bdf50)
            ins.append(("",
                f"When scoring 50+, team wins <strong>{wp50}%</strong> "
                f"of matches — clear positive impact on team results", "pos"))

    if "Innings_no" in bdf.columns and "Not_out" in bdf.columns:
        chase_no = bdf[(bdf["Innings_no"] == 2) & (bdf["Not_out"] == 1)]
        if not chase_no.empty:
            wp_no = win_pct(chase_no)
            ins.append(("",
                f"When remaining not out in chases, team wins "
                f"<strong>{wp_no}%</strong> of matches — "
                f"outstanding finishing ability", "pos"))

    bnd_pct = safe_pct(bnd, max(runs, 1))
    ins.append(("",
        f"Boundary dependency: <strong>{bnd_pct}%</strong> — "
        f"{'heavy reliance on boundaries' if bnd_pct > 50 else 'balanced attack across the ground'}",
        "neg" if bnd_pct > 50 else "pos"))

    return ins


def key_insights_bowl(bdf: pd.DataFrame) -> list:
    if bdf.empty:
        return []
    ins = []
    bs  = bowl_stats(bdf)

    ins.append(("",
        f"Bowling strike rate: <strong>{bs['bsr']}</strong> balls per wicket "
        f"| Dot ball %: <strong>{bs['dp']}%</strong>",
        "pos" if bs["bsr"] < 25 else "neu"))

    if "Match_result" in bdf.columns:
        w3 = bdf[bdf["Wickets"] >= 3]
        if not w3.empty:
            wp3 = win_pct(w3)
            ins.append(("",
                f"Win % when taking 3+ wickets: <strong>{wp3}%</strong> — "
                f"significant match-winner", "pos"))

    return ins


def recent_form_sentences(raw_df: pd.DataFrame, role: str) -> list:
    last10 = raw_df.sort_values("Match_date").tail(10)
    sents  = []
    rl     = role.lower()
    if "batsman" in rl or "all" in rl:
        b10 = bat_rows(last10)
        if not b10.empty:
            r5  = int(b10.tail(5)["Runs"].sum())
            sents.append(
                f"Scored <strong>{r5} runs</strong> in the last 5 matches.")
            sd = int((b10["Runs"] < 10).sum())
            if sd >= 3:
                sents.append(
                    f"Recorded <strong>{sd} single-digit scores</strong> "
                    f"in the last 10 matches — form concern.")
            no3 = int(b10.tail(3)["Not_out"].sum()) if "Not_out" in b10.columns else 0
            if no3 >= 2:
                sents.append(
                    f"Remained <strong>not out in {no3} of the last 3 innings</strong>"
                    f" — excellent finishing ability.")
            avg5 = calc_avg(b10.tail(5)["Runs"],
                            b10.tail(5).get("Not_out",
                                pd.Series(0, index=b10.tail(5).index)),
                            len(b10.tail(5)))
            if avg5 >= 40:
                sents.append(
                    f"Maintaining outstanding consistency — "
                    f"last 5 match average: <strong>{avg5}</strong>.")
            elif avg5 < 20:
                sents.append(
                    f"Recent form dip detected — "
                    f"last 5 match average: <strong>{avg5}</strong>.")

    if "bowler" in rl or "all" in rl:
        bw10 = bowl_rows(last10)
        if not bw10.empty:
            w5  = int(bw10.tail(5)["Wickets"].sum())
            ec5 = round(bw10.tail(5)["Economy"].mean(), 2)
            sents.append(
                f"Claimed <strong>{w5} wickets</strong> in the last 5 matches.")
            sents.append(
                f"Recent economy rate: <strong>{ec5} RPO</strong> "
                f"over the last 5 outings.")

    return sents[:4]

# ═══════════════════════════════════════════════════════════════
#  HTML BUILDERS
# ═══════════════════════════════════════════════════════════════
def kpi(val, lbl, variant=""):
    return (f'<div class="kpi-c {variant}">'
            f'<div class="kpi-v">{val}</div>'
            f'<div class="kpi-l">{lbl}</div></div>')

def ins_html(icon, txt, typ):
    cls = {"pos":"ins-pos","neg":"ins-neg","neu":"ins-neu"}.get(typ,"ins-neu")
    return (f'<div class="ins {cls}">'
            f'<span class="ins-icon">{icon}</span>'
            f'<span class="ins-body">{txt}</span></div>')

def sec_head(num, icon, title):
    return (f'<div class="rpt-sec-head">'
            f'<span class="sec-num">{num}</span>'
            f'<span class="sec-label">{icon} &nbsp;{title}</span>'
            f'</div>')

def subsec(title):
    return f'<div class="rpt-subsec">{title}</div>'

# ═══════════════════════════════════════════════════════════════
#  NAVBAR
# ═══════════════════════════════════════════════════════════════
st.markdown("""
<div class="navbar">
  <div class="nav-brand">
     Cricket Player <span>Performance Analysis</span>
  </div>
  <div class="nav-right">
    <a class="nb nb-solid" href="/" target="_self">Back</a>
    <a class="nb nb-solid" href="/userhome" target="_self">Home</a>
    <a class="nb nb-solid" href="/admin" target="_self">Admin</a>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ptb">
  <h2>Performance Report Generator</h2>
  <p>Generate a comprehensive professional cricket analytics report — export as PDF</p>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
#  FILTER PANEL
# ═══════════════════════════════════════════════════════════════
default_player = st.session_state.get(
    "rpt_player",
    st.session_state.get("cmp_player",
    st.session_state.get("ow_player",
    st.session_state.get("op_player",
    sorted(stats_dict.keys())[0]))))

fc1, fc2, fc3, fc4 = st.columns([3, 2, 2, 2])

all_players = sorted(stats_dict.keys())
with fc1:
    st.markdown('<div class="fl">Player</div>', unsafe_allow_html=True)
    p_idx      = all_players.index(default_player) if default_player in all_players else 0
    sel_player = st.selectbox("P", all_players, index=p_idx,
                              label_visibility="collapsed")

_raw = stats_dict.get(sel_player, pd.DataFrame())

with fc2:
    st.markdown('<div class="fl">Opponent</div>', unsafe_allow_html=True)
    opps   = (sorted(_raw["Opponent"].dropna().unique())
              if not _raw.empty and "Opponent" in _raw.columns else [])
    sel_opp = st.selectbox("O", ["All"] + opps, label_visibility="collapsed")

with fc3:
    st.markdown('<div class="fl">Country</div>', unsafe_allow_html=True)
    ctrs   = (sorted(_raw["Country"].dropna().unique())
              if not _raw.empty and "Country" in _raw.columns else [])
    sel_ctr = st.selectbox("C", ["All"] + ctrs, label_visibility="collapsed")

with fc4:
    st.markdown('<div class="fl">Year</div>', unsafe_allow_html=True)
    yrs    = (sorted(_raw["Match_date"].dt.year.dropna().astype(int).unique())
              if not _raw.empty else [])
    sel_year = st.selectbox("Y", ["All"] + [str(y) for y in yrs],
                            label_visibility="collapsed")

gen_btn = st.button("GENERATE", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)

if gen_btn:
    st.session_state.update(rpt_ok=True, rpt_player=sel_player,
                            rpt_opp=sel_opp, rpt_ctr=sel_ctr,
                            rpt_year=sel_year)

if not st.session_state.get("rpt_ok"):
    st.markdown(
        '<div class="nodata"><span class="nd-icon"></span>'
        '<p>Select a player &amp; filters, then click <strong>Generate</strong>.</p>'
        '</div>', unsafe_allow_html=True)
    st.stop()

sel_player = st.session_state["rpt_player"]
sel_opp    = st.session_state["rpt_opp"]
sel_ctr    = st.session_state["rpt_ctr"]
sel_year   = st.session_state["rpt_year"]

prow    = master_df[master_df["PlayerName"].str.strip() == sel_player]
pm      = prow.iloc[0] if not prow.empty else None
p_role  = str(pm["Role"])          if pm is not None else "Batsman"
p_ctr   = str(pm["Country"])       if pm is not None else "India"
p_bat   = str(pm["BattingStyle"])  if pm is not None else "RHB"
p_bowl  = str(pm["BowlingStyle"])  if pm is not None else "—"
p_dob   = str(pm["DOB"])           if pm is not None else "N/A"
p_debut = str(pm["DebutYear"])     if pm is not None else "N/A"
p_pid   = str(pm["PlayerID"])      if pm is not None else "—"
p_img   = str(pm["ImageFileName"]) if pm is not None else ""

rl       = role_norm(p_role)
is_bat   = rl == "bat"
is_bowl  = rl == "bowl"
is_ar    = rl == "ar"

wdf = _raw.copy()
if sel_opp  != "All": wdf = wdf[wdf["Opponent"] == sel_opp]
if sel_ctr  != "All": wdf = wdf[wdf["Country"]  == sel_ctr]
if sel_year != "All": wdf = wdf[wdf["Match_date"].dt.year == int(sel_year)]
wdf = wdf.sort_values("Match_date")

bat_df  = bat_rows(wdf)
bowl_df = bowl_rows(wdf)

gen_ts   = datetime.now().strftime("%d %B %Y,  %H:%M")
gen_date = datetime.now().strftime("%d %B %Y")
rpt_id   = f"RPT-{p_pid}-{datetime.now().strftime('%Y%m%d%H%M')}"

scope_parts = []
if sel_opp  != "All": scope_parts.append(f"vs {sel_opp}")
if sel_ctr  != "All": scope_parts.append(f"in {sel_ctr}")
if sel_year != "All": scope_parts.append(str(sel_year))
scope_lbl = "  .  ".join(scope_parts) if scope_parts else "All Matches"

if is_bat:
    perf_ins = batsman_insights(bat_df, _raw)
    key_ins  = key_insights_bat(bat_df)
elif is_bowl:
    perf_ins = bowler_insights(bowl_df)
    key_ins  = key_insights_bowl(bowl_df)
else:
    perf_ins = batsman_insights(bat_df, _raw) + bowler_insights(bowl_df)
    key_ins  = key_insights_bat(bat_df) + key_insights_bowl(bowl_df)

rf_sents = recent_form_sentences(_raw, p_role)

# ═══════════════════════════════════════════════════════════════
#  REPORT PREVIEW 
# ═══════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="masthead">
  <div class="masthead-left">
    <div class="masthead-text">
      <div class="mt">Cricket Player Performance Analysis</div>
      <div class="ms">Advanced Analytics Report &nbsp;·&nbsp;
        {sel_player} &nbsp;·&nbsp; {scope_lbl}</div>
    </div>
  </div>
  <div class="masthead-right">
    <div class="mr-date">Generated: {gen_ts}</div>
    <div class="mr-id">{rpt_id}</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown(sec_head("1","","Player Profile"), unsafe_allow_html=True)

img_src  = get_img_b64(p_img) if p_img else None
img_html = (f'<img src="{img_src}" style="width:110px;height:110px;'
            f'border-radius:50%;object-fit:cover;border:3px solid #1B3A6B;'
            f'display:block;margin:0 auto;">'
            if img_src else avatar_html(sel_player))

rb_cls = role_badge_cls(p_role)
chip_html = (
    f'<span class="chip">Opponent: {sel_opp}</span>'
    f'<span class="chip">Country: {sel_ctr}</span>'
    f'<span class="chip">Year: {sel_year}</span>'
    f'<span class="chip">Scope: {scope_lbl}</span>'
)

st.markdown(f"""
<div class="profile-outer">
  <div class="profile-img-col">
    {img_html}
    <div class="pname">{sel_player}</div>
    <div style="text-align:center;">
      <span class="role-badge {rb_cls}">{p_role}</span>
    </div>
  </div>
  <div class="profile-info-col">
    <div class="info-grid">
      <div>
        <div class="info-row"><span class="il">Country</span><span class="iv">{p_ctr}</span></div>
        <div class="info-row"><span class="il">Batting Style</span><span class="iv">{p_bat}</span></div>
        <div class="info-row"><span class="il">Bowling Style</span><span class="iv">{p_bowl}</span></div>
      </div>
      <div>
        <div class="info-row"><span class="il">Date of Birth</span><span class="iv">{p_dob}</span></div>
        <div class="info-row"><span class="il">Debut Year</span><span class="iv">{p_debut}</span></div>
        <div class="info-row"><span class="il">Player ID</span><span class="iv">{p_pid}</span></div>
      </div>
    </div>
    <div class="filter-chips-box">
      <div class="fct">Report Filters Applied</div>
      {chip_html}
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec_head("2","","Career Snapshot"), unsafe_allow_html=True)

if is_bat:
    bs = bat_stats(bat_df)
    if bs:
        st.markdown(
            f'<div class="kpi-band">'
            f'{kpi(bs["inn"],"Innings","")}'
            f'{kpi(bs["runs"],"Total Runs","c-amber")}'
            f'{kpi(bs["avg"],"Batting Avg","")}'
            f'{kpi(bs["sr"],"Strike Rate","")}'
            f'{kpi(bs["best"],"Best Score","c-amber")}'
            f'{kpi(str(bs["f4"])+"/"+str(bs["f6"]),"4s / 6s","")}'
            f'{kpi(str(bs["fif"])+"/"+str(bs["hun"]),"50s / 100s","c-green")}'
            f'{kpi(bs["no"],"Not Outs","")}'
            f'{kpi(str(bs["wp"])+"%","Win %","c-green")}'
            f'</div>', unsafe_allow_html=True)
elif is_bowl:
    bws = bowl_stats(bowl_df)
    if bws:
        st.markdown(
            f'<div class="kpi-band">'
            f'{kpi(bws["inn"],"Matches","")}'
            f'{kpi(bws["ovs"],"Overs","")}'
            f'{kpi(bws["mai"],"Maidens","")}'
            f'{kpi(bws["rc"],"Runs Conceded","c-amber")}'
            f'{kpi(bws["twk"],"Wickets","c-amber")}'
            f'{kpi(bws["eco"],"Economy","")}'
            f'{kpi(str(bws["dp"])+"%","Dot Ball %","c-green")}'
            f'{kpi(bws["bsr"],"Bowl SR","")}'
            f'{kpi(bws["bfig"],"Best Figures","c-green")}'
            f'{kpi(str(bws["wp"])+"%","Win %","")}'
            f'</div>', unsafe_allow_html=True)
else:
    st.markdown(subsec("Batting Performance"), unsafe_allow_html=True)
    bs = bat_stats(bat_df)
    if bs:
        st.markdown(
            f'<div class="kpi-band">'
            f'{kpi(bs["inn"],"Innings","")}'
            f'{kpi(bs["runs"],"Runs","c-amber")}'
            f'{kpi(bs["avg"],"Bat Avg","")}'
            f'{kpi(bs["sr"],"SR","")}'
            f'{kpi(str(bs["f4"])+"/"+str(bs["f6"]),"4s/6s","")}'
            f'{kpi(str(bs["fif"])+"/"+str(bs["hun"]),"50s/100s","c-green")}'
            f'{kpi(bs["best"],"Best","c-amber")}'
            f'</div>', unsafe_allow_html=True)
    st.markdown(subsec("Bowling Performance"), unsafe_allow_html=True)
    bws = bowl_stats(bowl_df)
    if bws:
        st.markdown(
            f'<div class="kpi-band">'
            f'{kpi(bws["ovs"],"Overs","")}'
            f'{kpi(bws["twk"],"Wickets","c-amber")}'
            f'{kpi(bws["eco"],"Economy","")}'
            f'{kpi(str(bws["dp"])+"%","Dot %","c-green")}'
            f'{kpi(bws["bsr"],"Bowl SR","")}'
            f'{kpi(bws["bfig"],"Best Fig","c-amber")}'
            f'</div>', unsafe_allow_html=True)

st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec_head("3","","Overall Performance Analysis"), unsafe_allow_html=True)
if perf_ins:
    blk = "".join(ins_html(icon, txt, typ) for icon, txt, typ in perf_ins[:12])
    st.markdown(blk, unsafe_allow_html=True)

st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec_head("4","","Recent Form Analysis (Last 10 Matches)"), unsafe_allow_html=True)
if rf_sents:
    blk = "".join(ins_html("", s, "neu") for s in rf_sents)
    st.markdown(blk, unsafe_allow_html=True)

last10 = _raw.sort_values("Match_date").tail(10)
if is_bat:
    lb10 = bat_rows(last10)
    if not lb10.empty:
        st.markdown(subsec("Last 10 Match Scorecard"), unsafe_allow_html=True)
        rows_h = ""
        for _, r in lb10.iterrows():
            res  = str(r.get("Match_result","—"))
            rc_  = "res-w" if res=="Won" else ("res-l" if res=="Lose" else "res-n")
            no_m = "*" if r.get("Not_out",0)==1 else ""
            rows_h += (f'<tr><td>{str(r["Match_date"])[:10]}</td>'
                       f'<td><strong>{r.get("Opponent","—")}</strong></td>'
                       f'<td>{r.get("Country","—")}</td>'
                       f'<td><strong>{int(r["Runs"])}{no_m}</strong></td>'
                       f'<td>{int(r.get("Balls_faced",0))}</td>'
                       f'<td>{r.get("Strike_Rate",0):.1f}</td>'
                       f'<td class="{rc_}">{res}</td></tr>')
        st.markdown(f"""<table class="ft"><thead><tr>
          <th>Date</th><th>Opponent</th><th>Country</th>
          <th>Runs</th><th>Balls</th><th>SR</th><th>Result</th>
        </tr></thead><tbody>{rows_h}</tbody></table>""", unsafe_allow_html=True)

elif is_bowl:
    lb10 = bowl_rows(last10)
    if not lb10.empty:
        st.markdown(subsec("Last 10 Match Figures"), unsafe_allow_html=True)
        rows_h = ""
        for _, r in lb10.iterrows():
            res  = str(r.get("Match_result","—"))
            rc_  = "res-w" if res=="Won" else ("res-l" if res=="Lose" else "res-n")
            tbl_ = int(r.get("Balls_Bowled",0))
            ov_s = f"{tbl_//6}.{tbl_%6}"
            rows_h += (f'<tr><td>{str(r["Match_date"])[:10]}</td>'
                       f'<td><strong>{r.get("Opponent","—")}</strong></td>'
                       f'<td>{r.get("Country","—")}</td>'
                       f'<td>{ov_s}</td>'
                       f'<td><strong>{int(r.get("Wickets",0))}</strong></td>'
                       f'<td>{int(r.get("Runs_Conceded",0))}</td>'
                       f'<td>{r.get("Economy",0):.2f}</td>'
                       f'<td class="{rc_}">{res}</td></tr>')
        st.markdown(f"""<table class="ft"><thead><tr>
          <th>Date</th><th>Opponent</th><th>Country</th>
          <th>Overs</th><th>Wkts</th><th>Runs</th>
          <th>Eco</th><th>Result</th>
        </tr></thead><tbody>{rows_h}</tbody></table>""", unsafe_allow_html=True)

st.markdown('<hr class="hr">', unsafe_allow_html=True)
st.markdown(sec_head("5","","Key Statistical Insights"), unsafe_allow_html=True)
if key_ins:
    blk = "".join(ins_html(icon, txt, typ) for icon, txt, typ in key_ins)
    st.markdown(blk, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════
#  PDF GENERATION
# ═══════════════════════════════════════════════════════════════
def build_pdf() -> bytes:
    buf = io.BytesIO()

    C_NAVY  = colors.HexColor("#1B3A6B")
    C_AMBER = colors.HexColor("#E8A020")
    C_STEEL = colors.HexColor("#4A6FA5")
    C_GREEN = colors.HexColor("#1E7A34")
    C_RED   = colors.HexColor("#C0392B")
    C_BG    = colors.HexColor("#f7f9fc")
    C_BORD  = colors.HexColor("#dde3ee")
    C_LIGHT = colors.HexColor("#EBF3FF")
    C_WHITE = colors.white

    # ── Page border + footer callback ───────────────────────
    def on_page(canvas, doc):
        W, H = A4
        canvas.saveState()
        canvas.setStrokeColor(C_NAVY)
        canvas.setLineWidth(1.2)
        canvas.rect(1.2*cm, 1.2*cm, W - 2.4*cm, H - 2.4*cm)
        canvas.setFont("Times-Roman", 8)
        canvas.setFillColor(C_STEEL)
        canvas.drawRightString(
            W - 1.5*cm, 1.0*cm,
            f"Page {doc.page}  .  {sel_player}  .  {scope_lbl}")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2.0*cm, rightMargin=2.0*cm,
        topMargin=2.0*cm,  bottomMargin=2.2*cm,
        title=f"Performance Report — {sel_player}",
    )

    # ── Style helpers ────────────────────────────────────────
    def PS(name, **kw):
        from reportlab.lib.styles import getSampleStyleSheet
        p = getSampleStyleSheet()["Normal"]
        return ParagraphStyle(name, parent=p, **kw)

    S_NORM  = PS("N",  fontName="Times-Roman", fontSize=10, leading=15)
    S_BOLD  = PS("B",  fontName="Times-Bold",  fontSize=10, leading=15)
    S_SM    = PS("SM", fontName="Times-Roman", fontSize=9,
                 textColor=C_STEEL, leading=13)
    S_WHITE = PS("WH", fontName="Times-Bold",  fontSize=15,
                 textColor=C_WHITE, leading=22)
    S_META  = PS("ME", fontName="Times-Roman", fontSize=9,
                 textColor=colors.HexColor("#bbbbbb"),
                 alignment=TA_RIGHT, leading=13)
    S_SEC   = PS("SC", fontName="Times-Bold",  fontSize=11,
                 textColor=C_NAVY, spaceBefore=10, spaceAfter=3, leading=16)
    S_SUB   = PS("SB", fontName="Times-Bold",  fontSize=10,
                 textColor=C_STEEL, spaceBefore=6, spaceAfter=3, leading=15)
    S_POS   = PS("IP", fontName="Times-Roman", fontSize=10,
                 textColor=C_GREEN, leftIndent=6, spaceAfter=2, leading=14)
    S_NEG   = PS("IN", fontName="Times-Roman", fontSize=10,
                 textColor=C_RED,   leftIndent=6, spaceAfter=2, leading=14)
    S_NEU   = PS("IU", fontName="Times-Roman", fontSize=10,
                 textColor=colors.HexColor("#6B4B00"),
                 leftIndent=6, spaceAfter=2, leading=14)
    S_KI    = PS("KI", fontName="Times-Roman", fontSize=10,
                 leftIndent=8, spaceAfter=3, leading=14)

    def section_heading(title):
        """Bold heading + navy underline, no number prefix."""
        return [
            Paragraph(title, S_SEC),
            HRFlowable(width="100%", thickness=0.8,
                       color=C_NAVY, spaceAfter=4),
        ]

    def sub_heading(title):
        return [Paragraph(title, S_SUB),
                HRFlowable(width="100%", thickness=0.4,
                           color=C_BORD, spaceAfter=3)]

    def hr():
        return HRFlowable(width="100%", thickness=0.4,
                          color=colors.HexColor("#cccccc"),
                          spaceAfter=6, spaceBefore=8)

    def insight_para(icon, txt, typ):
        clean = strip_tags(txt)
        s = S_POS if typ == "pos" else (S_NEG if typ == "neg" else S_NEU)
        pfx = "+" if typ == "pos" else ("-" if typ == "neg" else "*")
        return Paragraph(f"{pfx}  {clean}", s)

    def stat_table(metrics, p1_name, p2_name=None):
        """
        If p2_name is None: 2-col table (Label | Value).
        If p2_name given:  3-col table (Label | P1 | P2) with green/red.
        """
        C_N = C_NAVY; C_B = C_BORD

        if p2_name is None:
            hdr = [Paragraph("Metric", PS("TH0", fontName="Times-Bold",
                               fontSize=10, textColor=C_WHITE)),
                   Paragraph("Value",  PS("TH1", fontName="Times-Bold",
                               fontSize=10, textColor=C_WHITE,
                               alignment=TA_CENTER))]
            rows = [hdr]
            for lbl, val in metrics:
                rows.append([
                    Paragraph(str(lbl), S_SM),
                    Paragraph(str(val), PS(f"V{id(val)}", fontName="Times-Bold",
                               fontSize=10, textColor=C_NAVY,
                               alignment=TA_CENTER, leading=14)),
                ])
            t = Table(rows, colWidths=["55%", "45%"])
        else:
            hdr = [
                Paragraph("Metric", PS("TH0", fontName="Times-Bold",
                           fontSize=10, textColor=C_WHITE)),
                Paragraph(p1_name, PS("TH1", fontName="Times-Bold",
                           fontSize=10, textColor=C_WHITE,
                           alignment=TA_CENTER)),
                Paragraph(p2_name, PS("TH2", fontName="Times-Bold",
                           fontSize=10, textColor=C_WHITE,
                           alignment=TA_CENTER)),
            ]
            rows = [hdr]
            for lbl, v1, v2, hb in metrics:
                try:
                    n1 = float(str(v1).replace("*","").replace("%",""))
                    n2 = float(str(v2).replace("*","").replace("%",""))
                    if n1 == n2: c1 = c2 = C_NAVY
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
                rows.append([Paragraph(str(lbl), S_SM), cell(v1, c1), cell(v2, c2)])
            t = Table(rows, colWidths=["46%", "27%", "27%"])

        t.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), C_NAVY),
            ("ROWBACKGROUNDS",(0, 1), (-1,-1), [C_WHITE, C_BG]),
            ("LINEBELOW",     (0, 0), (-1,-1), 0.3, C_BORD),
            ("BOX",           (0, 0), (-1,-1), 0.5, C_BORD),
            ("TOPPADDING",    (0, 0), (-1,-1), 4),
            ("BOTTOMPADDING", (0, 0), (-1,-1), 4),
            ("LEFTPADDING",   (0, 0), (-1,-1), 6),
        ]))
        return t

    def kpi_table(items, cols=5):
        """KPI grid — label on top, bold value below."""
        while len(items) % cols != 0:
            items.append(("", ""))
        rows_ = []
        for i in range(0, len(items), cols):
            chunk = items[i:i+cols]
            vr = [Paragraph(f"<b>{v}</b>",
                            PS(f"KV{i+j}", fontName="Times-Bold", fontSize=12,
                               textColor=C_NAVY, alignment=TA_CENTER, leading=16))
                  for j, (_, v) in enumerate(chunk)]
            lr = [Paragraph(str(k), PS(f"KL{i+j}", fontName="Times-Roman",
                               fontSize=8, textColor=C_STEEL,
                               alignment=TA_CENTER, leading=12))
                  for j, (k, _) in enumerate(chunk)]
            rows_ += [vr, lr]
        cw = [f"{100/cols:.1f}%"] * cols
        t = Table(rows_, colWidths=cw)
        t.setStyle(TableStyle([
            ("ALIGN",        (0, 0), (-1,-1), "CENTER"),
            ("VALIGN",       (0, 0), (-1,-1), "MIDDLE"),
            ("ROWBACKGROUNDS",(0, 0), (-1,-1), [C_WHITE, C_BG] * 20),
            ("LINEBELOW",    (0, 0), (-1,-1), 0.3, C_BORD),
            ("TOPPADDING",   (0, 0), (-1,-1), 4),
            ("BOTTOMPADDING",(0, 0), (-1,-1), 4),
            ("BOX",          (0, 0), (-1,-1), 0.5, C_NAVY),
        ]))
        return t

    # ── BUILD STORY ─────────────────────────────────────────
    story = []

    # ── MASTHEAD (navy, keep as requested) ──────────────────
    mhd = Table([[
        Paragraph("Cricket Player Performance Analysis", S_WHITE),
        Paragraph(f"Report ID: {rpt_id}<br/>Generated: {gen_ts}", S_META),
    ]], colWidths=["65%", "35%"])
    mhd.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), C_NAVY),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 14),
        ("BOTTOMPADDING", (0,0),(-1,-1), 14),
        ("LEFTPADDING",   (0,0),(-1,-1), 14),
        ("RIGHTPADDING",  (0,0),(-1,-1), 14),
    ]))
    story += [mhd, Spacer(1, 8)]

    # Summary line under masthead
    story.append(Paragraph(
        f"<b>Player:</b> {sel_player}  &nbsp;&nbsp;  "
        f"<b>Role:</b> {p_role}  &nbsp;&nbsp;  "
        f"<b>Scope:</b> {scope_lbl}",
        PS("SUM", fontName="Times-Roman", fontSize=9, textColor=C_STEEL,
           leading=14, spaceAfter=8)))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=C_NAVY, spaceAfter=10))

    # ── 1. PLAYER PROFILE ───────────────────────────────────
    story += section_heading("Player Profile")
    story.append(Spacer(1, 6))

    # two-column profile table
    profile_data = [
        [Paragraph("Country",       S_SM), Paragraph(p_ctr,        S_BOLD)],
        [Paragraph("Batting Style", S_SM), Paragraph(p_bat,        S_BOLD)],
        [Paragraph("Bowling Style", S_SM), Paragraph(p_bowl,       S_BOLD)],
        [Paragraph("Date of Birth", S_SM), Paragraph(str(p_dob),   S_NORM)],
        [Paragraph("Debut Year",    S_SM), Paragraph(str(p_debut), S_NORM)],
        [Paragraph("Player ID",     S_SM), Paragraph(str(p_pid),   S_NORM)],
    ]
    prof_t = Table(profile_data, colWidths=["35%", "65%"])
    prof_t.setStyle(TableStyle([
        ("ROWBACKGROUNDS",(0,0),(-1,-1), [C_WHITE, C_BG]),
        ("LINEBELOW",    (0,0),(-1,-1), 0.3, C_BORD),
        ("TOPPADDING",   (0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("LEFTPADDING",  (0,0),(-1,-1), 8),
        ("BOX",          (0,0),(-1,-1), 0.5, C_BORD),
    ]))
    story.append(prof_t)
    story.append(Spacer(1, 4))

    # Filters applied
    story.append(Paragraph(
        f"Filters applied — Opponent: {sel_opp}  |  "
        f"Country: {sel_ctr}  |  Year: {sel_year}  |  Scope: {scope_lbl}",
        PS("FC", fontName="Times-Roman", fontSize=9,
           textColor=colors.HexColor("#8A6000"),
           leading=13, spaceAfter=4)))
    story.append(hr())

    # ── 2. CAREER SNAPSHOT ──────────────────────────────────
    story += section_heading("Career Snapshot")
    story.append(Spacer(1, 6))

    if is_bat:
        bs = bat_stats(bat_df)
        if bs:
            items = [
                ("Innings",    bs["inn"]),
                ("Total Runs", bs["runs"]),
                ("Batting Avg",bs["avg"]),
                ("Strike Rate",bs["sr"]),
                ("Best Score", bs["best"]),
                ("4s / 6s",   f"{bs['f4']}/{bs['f6']}"),
                ("50s / 100s",f"{bs['fif']}/{bs['hun']}"),
                ("Not Outs",   bs["no"]),
                ("Win %",     f"{bs['wp']}%"),
            ]
            story.append(kpi_table(items, cols=5))
        else:
            story.append(Paragraph("No batting data for these filters.", S_SM))

    elif is_bowl:
        bws = bowl_stats(bowl_df)
        if bws:
            items = [
                ("Matches",       bws["inn"]),
                ("Overs",         bws["ovs"]),
                ("Maidens",       bws["mai"]),
                ("Runs Conceded", bws["rc"]),
                ("Wickets",       bws["twk"]),
                ("Economy",       bws["eco"]),
                ("Dot Ball %",   f"{bws['dp']}%"),
                ("Bowl SR",       bws["bsr"]),
                ("Best Figures",  bws["bfig"]),
                ("Win %",        f"{bws['wp']}%"),
            ]
            story.append(kpi_table(items, cols=5))
        else:
            story.append(Paragraph("No bowling data for these filters.", S_SM))

    else:  # All-rounder
        story += sub_heading("Batting")
        bs = bat_stats(bat_df)
        if bs:
            items = [
                ("Innings",  bs["inn"]), ("Runs",  bs["runs"]),
                ("Bat Avg",  bs["avg"]), ("SR",    bs["sr"]),
                ("4s/6s",   f"{bs['f4']}/{bs['f6']}"),
                ("50s/100s",f"{bs['fif']}/{bs['hun']}"),
                ("Best",     bs["best"]),
            ]
            story.append(kpi_table(items, cols=4))
        story.append(Spacer(1, 6))
        story += sub_heading("Bowling")
        bws = bowl_stats(bowl_df)
        if bws:
            items = [
                ("Overs",  bws["ovs"]), ("Wickets", bws["twk"]),
                ("Economy",bws["eco"]), ("Dot %",  f"{bws['dp']}%"),
                ("Bowl SR", bws["bsr"]), ("Best Fig",bws["bfig"]),
            ]
            story.append(kpi_table(items, cols=4))

    story.append(hr())

    # ── 3. OVERALL PERFORMANCE ANALYSIS ─────────────────────
    story += section_heading("Overall Performance Analysis")
    story.append(Spacer(1, 4))
    for icon, txt, typ in perf_ins[:12]:
        story.append(insight_para(icon, txt, typ))
        story.append(Spacer(1, 2))
    story.append(hr())

    # ── 4. RECENT FORM ───────────────────────────────────────
    story += section_heading("Recent Form Analysis (Last 10 Matches)")
    story.append(Spacer(1, 4))
    for s in rf_sents:
        story.append(Paragraph(f"*  {strip_tags(s)}", S_NEU))
        story.append(Spacer(1, 2))

    # Scorecard table
    last10 = _raw.sort_values("Match_date").tail(10)
    if is_bat:
        lb10 = bat_rows(last10)
        if not lb10.empty:
            story.append(Spacer(1, 6))
            story += sub_heading("Last 10 Match Scorecard")
            hdr = ["Date", "Opponent", "Country", "Runs", "Balls", "SR", "Result"]
            tdata = [hdr]
            for _, r in lb10.iterrows():
                no_m = "*" if r.get("Not_out", 0) == 1 else ""
                tdata.append([
                    str(r["Match_date"])[:10],
                    str(r.get("Opponent", "—")),
                    str(r.get("Country", "—")),
                    f"{int(r['Runs'])}{no_m}",
                    str(int(r.get("Balls_faced", 0))),
                    f"{r.get('Strike_Rate', 0):.1f}",
                    str(r.get("Match_result", "—")),
                ])
            sc = Table(tdata, colWidths=["15%","18%","13%","10%","10%","10%","14%"])
            sc.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),(-1,0), C_NAVY),
                ("TEXTCOLOR",     (0,0),(-1,0), C_WHITE),
                ("FONTNAME",      (0,0),(-1,0), "Times-Bold"),
                ("FONTSIZE",      (0,0),(-1,-1), 9),
                ("ALIGN",         (0,0),(-1,-1), "CENTER"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1), [C_WHITE, C_BG]),
                ("LINEBELOW",     (0,0),(-1,-1), 0.3, C_BORD),
                ("TOPPADDING",    (0,0),(-1,-1), 4),
                ("BOTTOMPADDING", (0,0),(-1,-1), 4),
                ("BOX",           (0,0),(-1,-1), 0.5, C_BORD),
            ]))
            story.append(sc)

    elif is_bowl:
        lb10 = bowl_rows(last10)
        if not lb10.empty:
            story.append(Spacer(1, 6))
            story += sub_heading("Last 10 Match Figures")
            hdr = ["Date","Opponent","Country","Overs","Wkts","Runs","Eco","Result"]
            tdata = [hdr]
            for _, r in lb10.iterrows():
                tbl_ = int(r.get("Balls_Bowled", 0))
                tdata.append([
                    str(r["Match_date"])[:10],
                    str(r.get("Opponent", "—")),
                    str(r.get("Country", "—")),
                    f"{tbl_//6}.{tbl_%6}",
                    str(int(r.get("Wickets", 0))),
                    str(int(r.get("Runs_Conceded", 0))),
                    f"{r.get('Economy', 0):.2f}",
                    str(r.get("Match_result", "—")),
                ])
            sc = Table(tdata, colWidths=["13%","15%","12%","9%","9%","9%","9%","14%"])
            sc.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),(-1,0), C_NAVY),
                ("TEXTCOLOR",     (0,0),(-1,0), C_WHITE),
                ("FONTNAME",      (0,0),(-1,0), "Times-Bold"),
                ("FONTSIZE",      (0,0),(-1,-1), 9),
                ("ALIGN",         (0,0),(-1,-1), "CENTER"),
                ("ROWBACKGROUNDS",(0,1),(-1,-1), [C_WHITE, C_BG]),
                ("LINEBELOW",     (0,0),(-1,-1), 0.3, C_BORD),
                ("TOPPADDING",    (0,0),(-1,-1), 4),
                ("BOTTOMPADDING", (0,0),(-1,-1), 4),
                ("BOX",           (0,0),(-1,-1), 0.5, C_BORD),
            ]))
            story.append(sc)

    story.append(hr())

    # ── 5. KEY STATISTICAL INSIGHTS ─────────────────────────
    story += section_heading("Key Statistical Insights")
    story.append(Spacer(1, 4))
    for icon, txt, typ in key_ins:
        story.append(insight_para(icon, txt, typ))
        story.append(Spacer(1, 2))

    story.append(Spacer(1, 14))

    # ── FOOTER ──────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=0.8,
                             color=C_NAVY, spaceAfter=4))

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════════
#  DOWNLOAD BAR
# ═══════════════════════════════════════════════════════════════
st.markdown('<div class="dl-bar">', unsafe_allow_html=True)
dl_a, dl_b, dl_c = st.columns([3.5, 3, 1.5])

with dl_a:
    st.markdown(
        f'<div class="dl-info">'
        f'<strong>{sel_player}</strong> &nbsp;.&nbsp; {scope_lbl}'
        f'<br>Report ID: {rpt_id} &nbsp;.&nbsp; {gen_ts}'
        f'</div>', unsafe_allow_html=True)

with dl_b:
    st.markdown("", unsafe_allow_html=True)

with dl_c:
    pdf_bytes = build_pdf()
    fname = (f"{sel_player.replace(' ','_')}_Report_"
             f"{sel_year}_{datetime.now().strftime('%Y%m%d')}.pdf")
    st.download_button(
        label="Download PDF",
        data=pdf_bytes,
        file_name=fname,
        mime="application/pdf",
        use_container_width=True,
    )

st.markdown('</div>', unsafe_allow_html=True)

st.markdown("""
<div class="pg-foot">
  Cricket Player Performance Analysis Dashboard &nbsp;|&nbsp; Built with Streamlit
</div>
""", unsafe_allow_html=True)