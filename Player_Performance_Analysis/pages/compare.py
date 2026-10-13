import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl, base64, os

# ── PAGE CONFIG ───────────────────────────────────────────────
st.set_page_config(
    page_title="Compare Players | Cricket Analysis",
    layout="wide",
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

# ── BRAND PALETTE ─────────────────────────────────────────────
NAVY  = "#1B3A6B"
AMBER = "#E8A020"
STEEL = "#4A6FA5"
GREEN = "#1E7A34"
RED   = "#C0392B"
BORDER= "#dde3ee"

P1_CLR  = NAVY
P2_CLR  = AMBER
P1_FILL = "rgba(27,58,107,0.18)"
P2_FILL = "rgba(232,160,32,0.18)"

SENA_COUNTRIES = {"South Africa","England","New Zealand","Australia"}
ASIA_COUNTRIES = {"India","Sri Lanka","Bangladesh","Pakistan","UAE"}

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

/* ─ NAVBAR ─ */
.top-header{
    background:#fff;border-bottom:2px solid #1B3A6B;
    padding:0 60px;height:64px;
    display:flex;align-items:center;justify-content:space-between;
    position:fixed;width:100%;top:0;z-index:999;
    box-shadow:0 2px 12px rgba(27,58,107,.08);
    margin:0 -60px 0 -60px;
}
.header-brand{font-family:'Poppins',sans-serif;font-size:1.2rem;
    font-weight:700;color:#1B3A6B;}
.header-brand span{color:#E8A020;}
.header-nav{display:flex;gap:8px;align-items:center;}
.h-btn{background:#1B3A6B;color:#fff !important;border:none;
    padding:7px 18px;border-radius:6px;font-size:.82rem;font-weight:600;
    text-decoration:none !important;display:inline-block;transition:background .2s;}
.h-btn:hover{background:#243f7a;}
.h-btn-ghost{background:transparent;color:#1B3A6B !important;
    border:1.5px solid #1B3A6B;padding:6px 16px;border-radius:6px;
    font-size:.82rem;font-weight:600;text-decoration:none !important;
    display:inline-block;transition:background .2s;}
.h-btn-ghost:hover{background:#f0f4fb;}

/* ─ PAGE TITLE ─ */
.page-title-bar{background:#f7f9fc;border-left:4px solid #1B3A6B;
    padding:14px 60px;margin:40px -60px 28px -60px;}
.page-title-bar h2{font-family:'Poppins',sans-serif;font-size:1.15rem;
    font-weight:700;color:#1B3A6B;margin:0;}
.page-title-bar p{color:#6B7A8D;font-size:.8rem;margin:3px 0 0 0;}

/* ─ FILTER ─ */
.filter-wrap{background:#f7f9fc;border:1px solid #dde3ee;border-radius:10px;
    padding:22px 28px;margin-bottom:26px;}
.filter-label{font-size:.75rem;font-weight:600;color:#1B3A6B;
    letter-spacing:.8px;text-transform:uppercase;margin-bottom:6px;}
.stSelectbox>div>div{background:#fff !important;border:1.5px solid #c8d0e0 !important;
    border-radius:7px !important;color:#1B3A6B !important;font-size:.88rem !important;}
.stSelectbox>div>div:focus-within{border-color:#1B3A6B !important;
    box-shadow:0 0 0 3px rgba(27,58,107,.1) !important;}
div[data-testid="stButton"]>button{background:#1B3A6B !important;color:#fff !important;
    font-family:'Poppins',sans-serif !important;font-weight:700 !important;
    font-size:.9rem !important;letter-spacing:1px !important;border:none !important;
    border-radius:7px !important;padding:12px 0 !important;width:100% !important;}
div[data-testid="stButton"]>button:hover{background:#243f7a !important;}

/* ─ SECTION TITLE ─ */
.sec-title{font-family:'Poppins',sans-serif;font-size:1rem;font-weight:700;
    color:#1B3A6B;margin:0 0 14px 0;padding-bottom:8px;
    border-bottom:2px solid #e8edf5;display:flex;align-items:center;gap:8px;}
.sec-title .dot{width:8px;height:8px;background:#1B3A6B;
    border-radius:50%;display:inline-block;}

/* ─ PROFILE CARDS ─ */
.profile-card{background:#fff;border:1px solid #dde3ee;border-radius:12px;
    padding:28px 22px;box-shadow:0 2px 12px rgba(27,58,107,.07);
    position:relative;overflow:hidden;text-align:center;}
.p1-card::before{content:'';position:absolute;top:0;left:0;right:0;
    height:5px;background:#1B3A6B;}
.p2-card::before{content:'';position:absolute;top:0;left:0;right:0;
    height:5px;background:#E8A020;}
.player-name{font-family:'Poppins',sans-serif;font-size:1.15rem;
    font-weight:700;color:#1B3A6B;text-align:center;margin:12px 0 4px 0;}
.player-role-badge{display:block;width:fit-content;margin:0 auto 16px auto;
    padding:3px 14px;border-radius:20px;font-size:.72rem;font-weight:700;
    letter-spacing:.5px;text-transform:uppercase;}
.badge-batsman{background:#EBF3FF;color:#1B3A6B;border:1px solid #c0d0eb;}
.badge-bowler{background:#FFF0E8;color:#C0501A;border:1px solid #f0c0a0;}
.badge-allrounder{background:#F0F8EE;color:#2A7A2A;border:1px solid #a8d4a0;}
.info-row{display:flex;justify-content:space-between;align-items:center;
    padding:8px 0;border-bottom:1px solid #f0f3f8;font-size:.82rem;}
.info-row:last-child{border-bottom:none;}
.info-lbl{color:#6B7A8D;font-weight:500;}
.info-val{color:#1B3A6B;font-weight:600;}

/* VS badge */
.vs-badge{display:flex;flex-direction:column;align-items:center;
    justify-content:center;height:100%;}
.vs-circle{width:80px;height:80px;border-radius:50%;
    background:linear-gradient(135deg,#1B3A6B,#E8A020);
    display:flex;align-items:center;justify-content:center;
    font-family:'Poppins',sans-serif;font-size:1.5rem;font-weight:800;
    color:#fff;box-shadow:0 4px 20px rgba(27,58,107,.25);}
.vs-sub{font-size:.74rem;color:#8899aa;margin-top:10px;
    text-align:center;font-weight:500;}

/* MODE BANNER */
.mode-banner{background:linear-gradient(135deg,#1B3A6B,#243f7a);
    border-radius:10px;padding:14px 24px;
    display:flex;align-items:center;gap:16px;margin-bottom:24px;}
.mode-pill{background:rgba(255,255,255,.15);color:#fff;
    padding:4px 14px;border-radius:20px;font-size:.72rem;
    font-weight:700;letter-spacing:.5px;}
.mode-title{font-family:'Poppins',sans-serif;font-size:1rem;
    font-weight:700;color:#fff;}
.mode-sub{color:rgba(255,255,255,.65);font-size:.78rem;margin-top:2px;}

/* ─ COMPARISON TABLE — centered metric col ─ */
.cmp-table-wrap{
    border-radius:12px;border:1px solid #dde3ee;
    overflow:hidden;box-shadow:0 2px 10px rgba(27,58,107,.06);
}
.cmp-table{width:100%;border-collapse:collapse;font-size:.85rem;}
.cmp-table colgroup col.c-p1 {width:35%;}
.cmp-table colgroup col.c-metric {width:30%;}
.cmp-table colgroup col.c-p2 {width:35%;}

/* Header */
.cmp-table thead tr{background:#1B3A6B;}
.cmp-table thead th{
    padding:12px 16px;
    font-family:'Poppins',sans-serif;font-size:.75rem;font-weight:700;
    letter-spacing:.6px;text-transform:uppercase;
    color:#fff;text-align:center;
}
.cmp-table thead th.th-metric{text-align:center;background:#243f7a;}

/* Sub-header row (section label) */
.cmp-table .sub-hdr td{
    background:#EBF3FF;
    font-family:'Poppins',sans-serif;font-size:.72rem;font-weight:700;
    color:#1B3A6B;letter-spacing:.5px;text-transform:uppercase;
    padding:8px 16px;text-align:center;
    border-top:2px solid #c0d0eb;border-bottom:1px solid #dde3ee;
}

/* Body rows */
.cmp-table tbody tr{border-bottom:1px solid #f0f3f8;}
.cmp-table tbody tr:last-child{border-bottom:none;}
.cmp-table tbody tr:hover td{background:#fafbfd;}
.cmp-table tbody td{
    padding:11px 18px;color:#2c3e50;font-weight:500;
    text-align:center;
}
/* Metric cell — bold, centered */
.cmp-table tbody td.td-metric{
    text-align:center;
    color:#1B3A6B;font-weight:700;
    font-family:'Poppins',sans-serif;font-size:.78rem;
    background:#f7f9fc;
    border-left:1px solid #e8edf5;
    border-right:1px solid #e8edf5;
}
/* Best / worst highlights */
.val-best{
    background:#e8f5ec !important;
    color:#1E7A34 !important;
    font-family:'Poppins',sans-serif;font-weight:700;
}
.val-worst{
    background:#fde8e8 !important;
    color:#C0392B !important;
    font-family:'Poppins',sans-serif;font-weight:700;
}
.val-equal{color:#1B3A6B;font-family:'Poppins',sans-serif;font-weight:700;}

/* P1 / P2 header name badges */
.p1-badge{background:#EBF3FF;color:#1B3A6B;padding:2px 10px;
    border-radius:4px;font-size:.7rem;font-weight:700;}
.p2-badge{background:#FFF8EE;color:#C07010;padding:2px 10px;
    border-radius:4px;font-size:.7rem;font-weight:700;}

/* CHART CARD */
.chart-card{background:#fff;border:1px solid #dde3ee;border-radius:12px;
    padding:20px 18px 10px 18px;
    box-shadow:0 2px 10px rgba(27,58,107,.06);margin-bottom:20px;}

/* INFO PILL (comparison mode label) */
.cmp-mode-pill{
    display:inline-flex;align-items:center;gap:8px;
    background:#EBF3FF;border:1px solid #c0d0eb;
    border-radius:8px;padding:8px 16px;
    font-size:.78rem;color:#1B3A6B;font-weight:600;
    margin-bottom:18px;
}

/* ROLE MISMATCH WARNING */
.role-warn{background:#fff8e8;border:1.5px solid #E8A020;
    border-radius:10px;padding:18px 24px;
    display:flex;align-items:center;gap:14px;margin:20px 0;}
.role-warn-icon{font-size:1.6rem;}
.role-warn-text{font-size:.88rem;color:#7a5800;}
.role-warn-text strong{color:#1B3A6B;}

/* DIVIDER / NO DATA */
.divider{border:none;height:1px;background:#e8edf5;margin:28px 0;}
.no-data{text-align:center;padding:40px;color:#8899aa;font-size:.9rem;}
.no-data .icon{font-size:2.5rem;display:block;margin-bottom:10px;}

/* ACTION BAR + FOOTER */
.action-bar{background:#f7f9fc;border-top:1px solid #dde3ee;
    padding:26px 60px;margin:32px -60px 0 -60px;text-align:center;}
.act-btn{display:inline-flex;align-items:center;gap:8px;
    padding:12px 32px;border-radius:7px;
    font-family:'Poppins',sans-serif;font-weight:700;font-size:.9rem;
    letter-spacing:.5px;cursor:pointer;
    text-decoration:none !important;border:none;}
.act-report{background:#1B3A6B;color:#fff !important;}
.act-report:hover{background:#243f7a;}
.footer{background:#1B3A6B;color:rgba(255,255,255,.7);text-align:center;
    padding:18px 60px;font-size:.8rem;
    margin:0 -60px -60px -60px;letter-spacing:.3px;}
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

# ── HELPERS ───────────────────────────────────────────────────
ALL_PLAYERS = sorted(stats_dict.keys())

SENA_COUNTRIES = {"South Africa","England","New Zealand","Australia"}
ASIA_COUNTRIES = {"India","Sri Lanka","Bangladesh","Pakistan","UAE"}

def get_role(name):
    """Return role string from master, or auto-detect from columns."""
    row = master_df[master_df["PlayerName"].str.strip() == name]
    if not row.empty:
        return str(row.iloc[0].get("Role","Batsman"))
    # Auto-detect from data columns
    df = stats_dict.get(name, pd.DataFrame())
    has_bat  = "Batted_Flag" in df.columns
    has_bowl = "Bowl_Flag"   in df.columns
    if has_bat and has_bowl: return "All-Rounder"
    if has_bowl:             return "Bowler"
    return "Batsman"

def role_norm(r):
    r = r.lower()
    if "all" in r: return "all-rounder"
    if "bowl" in r: return "bowler"
    return "batsman"

def get_meta(name):
    row = master_df[master_df["PlayerName"].str.strip() == name]
    if not row.empty:
        p = row.iloc[0]
        return dict(country=str(p.get("Country","India")),
                    role=str(p.get("Role","Batsman")),
                    bat=str(p.get("BattingStyle","RHB")),
                    bowl=str(p.get("BowlingStyle","—")),
                    dob=str(p.get("DOB","N/A")),
                    debut=str(p.get("DebutYear","N/A")),
                    img=str(p.get("ImageFileName","")))
    return dict(country="India", role=get_role(name), bat="RHB",
                bowl="—", dob="N/A", debut="N/A", img="")

def get_img_b64(filename):
    for d in [IMAGE_DIR, os.path.join(ROOT_DIR,"images"), BASE_DIR]:
        p = os.path.join(d, filename)
        if os.path.exists(p):
            ext  = filename.rsplit(".",1)[-1].lower()
            mime = "jpeg" if ext in ("jpg","jpeg") else ext
            with open(p,"rb") as f:
                return f"data:image/{mime};base64,{base64.b64encode(f.read()).decode()}"
    return None

def initials_avatar(name, clr=NAVY):
    ini = "".join(p[0].upper() for p in name.split()[:2])
    bg  = "#EBF3FF" if clr == NAVY else "#FFF8E8"
    return (f'<div style="width:90px;height:90px;border-radius:50%;background:{bg};'
            f'border:3px solid {clr};display:flex;align-items:center;justify-content:center;'
            f'margin:0 auto;font-family:Poppins,sans-serif;font-size:1.6rem;font-weight:700;'
            f'color:{clr};">{ini}</div>')

def calc_avg(runs_s, not_out_s, innings):
    tot = float(runs_s.sum())
    dis = max(int(innings) - int(not_out_s.sum()), 0)
    return round(tot/dis, 2) if dis > 0 else round(tot, 2)

def gcol(df, *names):
    for n in names:
        if n in df.columns: return df[n]
    return pd.Series(0, index=df.index)

def bat_rows(df): return df[df["Batted_Flag"]==1] if "Batted_Flag" in df.columns else pd.DataFrame()
def bowl_rows(df): return df[df["Bowl_Flag"]==1]   if "Bowl_Flag"   in df.columns else pd.DataFrame()

def badge_cls(role):
    r = role.lower()
    if "all" in r: return "badge-allrounder"
    if "bowl" in r: return "badge-bowler"
    return "badge-batsman"

def pchart(fig, h=320, title=""):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Inter",color="#2c3e50",size=11),
        height=h, margin=dict(l=10,r=10,t=46,b=10),
        title=dict(text=title,
            font=dict(family="Poppins",size=13,color=NAVY),
            x=0.02, xanchor="left"),
        legend=dict(bgcolor="rgba(255,255,255,.9)",
            bordercolor=BORDER,borderwidth=1,font=dict(size=10)),
    )
    fig.update_xaxes(gridcolor="#f0f3f8",linecolor=BORDER,
        tickfont=dict(size=9),zeroline=False)
    fig.update_yaxes(gridcolor="#f0f3f8",linecolor=BORDER,
        tickfont=dict(size=9),zeroline=False)
    return fig

def pch(fig):
    st.plotly_chart(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════════
#  COMPARISON LOGIC
# ════════════════════════════════════════════════════════════════
def compare_mode_label(r1, r2):

    # normalise
    n1 = role_norm(r1)
    n2 = role_norm(r2)

    if n1 == "batsman"     and n2 == "batsman":     return "batting", "Batting Stats Comparison"
    if n1 == "bowler"      and n2 == "bowler":      return "bowling", "Bowling Stats Comparison"
    if n1 == "all-rounder" and n2 == "all-rounder": return "both",    "Batting & Bowling Stats Comparison"

    # Batsman vs All-Rounder → batting only
    if (n1 == "batsman" and n2 == "all-rounder") or \
       (n1 == "all-rounder" and n2 == "batsman"):
        return "batting", "Batting Stats Comparison (Batsman vs All-Rounder)"

    # Bowler vs All-Rounder → bowling only
    if (n1 == "bowler" and n2 == "all-rounder") or \
       (n1 == "all-rounder" and n2 == "bowler"):
        return "bowling", "Bowling Stats Comparison (Bowler vs All-Rounder)"

    return None, (f"Cannot compare a {r1} with a {r2}. "
                  f"Valid pairings: Batsman↔Batsman, Batsman↔All-Rounder, "
                  f"Bowler↔Bowler, Bowler↔All-Rounder, All-Rounder↔All-Rounder.")

# ── Build filtered dataframe ───────────────────────────────────
def build_df(player, mode, sel_opp, sel_ctr, sel_year):
    """
    Correct filter pipeline — every active filter is always applied:

    Step 1  Mode-based primary filter (country group / specific country /
            specific opponent)
    Step 2  Year filter — applied on top whenever sel_year != "All"
    Step 3  Overall + specific opponent cross-filter
    Step 4  Recent Form — last 5 rows AFTER all other filters are done

    Examples:
      Opponent Wise + West Indies        -> only West Indies rows
      Country Wise + India + Year 2020   -> India rows, 2020 only
      SENA Countries                     -> SA/ENG/NZ/AUS rows
      Recent Form                        -> last 5 matches by date
      Recent Form + Year 2019            -> last 5 matches inside 2019
      Overall + Opponent=West Indies     -> only West Indies rows
      Overall + Year 2020                -> only 2020 rows
    """
    df = stats_dict.get(player, pd.DataFrame()).copy()
    if df.empty:
        return df, pd.DataFrame(), pd.DataFrame()

    # Step 1 — mode-based primary filter
    if mode == "SENA Countries":
        if sel_ctr not in ("All SENA", "All", ""):
            df = df[df["Country"] == sel_ctr] if "Country" in df.columns else df
        else:
            df = df[df["Country"].isin(SENA_COUNTRIES)] if "Country" in df.columns else df

    elif mode == "Asian Countries":
        if sel_ctr not in ("All Asia", "All", ""):
            df = df[df["Country"] == sel_ctr] if "Country" in df.columns else df
        else:
            df = df[df["Country"].isin(ASIA_COUNTRIES)] if "Country" in df.columns else df

    elif mode == "Country Wise":
        if sel_ctr not in ("All", ""):
            df = df[df["Country"] == sel_ctr] if "Country" in df.columns else df

    elif mode == "Opponent Wise":
        if sel_opp not in ("All", ""):
            df = df[df["Opponent"] == sel_opp] if "Opponent" in df.columns else df

    # Step 2 — year filter (always honoured when set, for every mode)
    if sel_year != "All":
        df = df[df["Match_date"].dt.year == int(sel_year)]

    # Step 3 — Overall mode with specific opponent selected
    if mode == "Overall" and sel_opp not in ("All", ""):
        df = df[df["Opponent"] == sel_opp] if "Opponent" in df.columns else df

    # Step 4 — Recent Form: last 5 by date AFTER all filters above
    if mode == "Recent Form":
        df = df.sort_values("Match_date").tail(5)

    return df, bat_rows(df), bowl_rows(df)

# ── Stats calculators ──────────────────────────────────────────
def compute_bat(bat_df):
    if bat_df.empty:
        return {k:"—" for k in ["Innings","Runs","Balls Faced","Batting Average",
            "Strike Rate","Best Score","4s","6s","50s","100s",
            "Dot Ball %","Boundary %","Not Outs"]}
    inn  = len(bat_df)
    runs = int(bat_df["Runs"].sum())
    bf   = int(bat_df["Balls_faced"].sum()) if "Balls_faced" in bat_df.columns else 0
    sr   = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(), 1)
    no_s = bat_df.get("Not_out", pd.Series(0, index=bat_df.index))
    no   = int(no_s.sum())
    avg  = calc_avg(bat_df["Runs"], no_s, inn)
    if not bat_df.empty:
        bi  = bat_df["Runs"].idxmax()
        best= f"{int(bat_df.loc[bi,'Runs'])}{'*' if int(bat_df.loc[bi,'Not_out']) else ''}"
    else: best = "0"
    f4   = int(bat_df["4s"].sum())
    f6   = int(bat_df["6s"].sum())
    fif  = int(bat_df["50s"].sum()) if "50s" in bat_df.columns else 0
    hun  = int(bat_df["100s"].sum()) if "100s" in bat_df.columns else 0
    dots = int(bat_df["Dots"].sum()) if "Dots" in bat_df.columns else 0
    dp   = round(dots/max(bf,1)*100, 1)
    bnd  = f4*4 + f6*6
    bp   = round(bnd/max(runs,1)*100, 1)
    return {"Innings": inn, "Runs": runs, "Balls Faced": bf,
            "Batting Average": avg, "Strike Rate": sr,
            "Best Score": best, "4s": f4, "6s": f6,
            "50s": fif, "100s": hun,
            "Dot Ball %": dp, "Boundary %": bp, "Not Outs": no}

def compute_bowl(bowl_df):
    if bowl_df.empty:
        return {k:"—" for k in ["Innings","Wickets","Balls Bowled","Economy",
            "Bowling Average","Best Figures","Bowling SR",
            "Dot Ball %","Maidens","Extras","RHB Wickets","LHB Wickets"]}
    inn  = len(bowl_df)
    tbl  = int(bowl_df["Balls_Bowled"].sum()) if "Balls_Bowled" in bowl_df.columns else 0
    twk  = int(bowl_df["Wickets"].sum())
    rc   = int(bowl_df["Runs_Conceded"].sum()) if "Runs_Conceded" in bowl_df.columns else 0
    eco  = round(bowl_df["Economy"].replace(0,float("nan")).mean(), 2)
    ba   = round(rc/max(twk,1), 2)
    bsr  = round(tbl/max(twk,1), 1)
    dots = int(gcol(bowl_df,"Dot_Balls").sum())
    dp   = round(dots/max(tbl,1)*100, 1)
    ext  = int(gcol(bowl_df,"Extras").sum())
    mai  = int(bowl_df["Maidens"].sum()) if "Maidens" in bowl_df.columns else 0
    rhbw = int(bowl_df["Wicket_RHB"].sum()) if "Wicket_RHB" in bowl_df.columns else 0
    lhbw = int(bowl_df["Wicket_LHB"].sum()) if "Wicket_LHB" in bowl_df.columns else 0
    if twk > 0:
        br  = bowl_df.sort_values(["Wickets","Runs_Conceded"],ascending=[False,True]).iloc[0]
        bfig= f"{int(br['Wickets'])}/{int(br.get('Runs_Conceded',0))}"
    else: bfig = "0/0"
    return {"Innings": inn, "Wickets": twk, "Balls Bowled": tbl,
            "Economy": eco, "Bowling Average": ba,
            "Best Figures": bfig, "Bowling SR": bsr,
            "Dot Ball %": dp, "Maidens": mai, "Extras": ext,
            "RHB Wickets": rhbw, "LHB Wickets": lhbw}

# Higher-is-better keys
HIGHER_BETTER = {"Innings","Runs","Balls Faced","Balls Bowled",
                 "Strike Rate","Batting Average","Best Score","4s","6s",
                 "50s","100s","Not Outs","Wickets","Dot Balls",
                 "Dot Ball %","RHB Wickets","LHB Wickets","Boundary %",
                 "Maidens"}
LOWER_BETTER  = {"Economy","Bowling Average","Extras","Bowling SR"}
SKIP_CMP      = {"Best Score","Best Figures"}

def cmp_classes(metric, v1, v2):
    if metric in SKIP_CMP: return "", ""
    try:
        n1 = float(str(v1).replace("*",""))
        n2 = float(str(v2).replace("*",""))
    except: return "", ""
    if n1 == n2: return "val-equal", "val-equal"
    if metric in LOWER_BETTER:
        return ("val-best","val-worst") if n1 < n2 else ("val-worst","val-best")
    return ("val-best","val-worst") if n1 > n2 else ("val-worst","val-best")

def build_table(s1, s2, p1_name, p2_name, section_label=None):
    """
    Returns HTML string for the comparison table.
    Layout: P1 value | Metric (center) | P2 value
    """
    header = f"""
    <div class="cmp-table-wrap">
    <table class="cmp-table">
      <colgroup>
        <col class="c-p1">
        <col class="c-metric">
        <col class="c-p2">
      </colgroup>
      <thead>
        <tr>
          <th><span class="p1-badge">{p1_name}</span></th>
          <th class="th-metric">Metric</th>
          <th><span class="p2-badge">{p2_name}</span></th>
        </tr>
      </thead>
      <tbody>
    """
    if section_label:
        header += f'<tr class="sub-hdr"><td colspan="3">{section_label}</td></tr>'

    rows_html = ""
    for metric in s1.keys():
        v1, v2 = s1.get(metric,"—"), s2.get(metric,"—")
        c1, c2 = cmp_classes(metric, v1, v2)
        rows_html += (
            f"<tr>"
            f"<td class='{c1}'>{v1}</td>"
            f"<td class='td-metric'>{metric}</td>"
            f"<td class='{c2}'>{v2}</td>"
            f"</tr>"
        )

    return header + rows_html + "</tbody></table></div>"

# ══════════════════════════════════════════════════════════════
#  NAVBAR
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div class="top-header">
  <div class="header-brand">Cricket Player<span> Performance Analysis</span></div>
  <div class="header-nav">
    <a class="h-btn" href="/userhome" target="_self">Home</a>
    <a class="h-btn" href="/"         target="_self">Back</a>
    <a class="h-btn"       href="/admin"    target="_self">Admin</a>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="page-title-bar">
  <h2>Player Comparison Analysis</h2>
  <p>Head-to-head comparison — Batsman vs Batsman / Batsman vs All-Rounder /
     Bowler vs Bowler / Bowler vs All-Rounder / All-Rounder vs All-Rounder</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SELECTION PANEL
# ══════════════════════════════════════════════════════════════
# Row 1 — Player selection
pc1, pvs, pc2 = st.columns([5, 1, 5])
default_p1 = st.session_state.get("cmp_player", ALL_PLAYERS[0])
default_p1 = default_p1 if default_p1 in ALL_PLAYERS else ALL_PLAYERS[0]

with pc1:
    st.markdown('<div class="filter-label">Player 1</div>', unsafe_allow_html=True)
    p1_idx = ALL_PLAYERS.index(default_p1)
    sel_p1 = st.selectbox("P1", ALL_PLAYERS, index=p1_idx,
                          label_visibility="collapsed", key="sel_p1")
with pvs:
    st.markdown('<div style="text-align:center;padding-top:28px;font-family:Poppins,'
                'sans-serif;font-size:1rem;font-weight:800;color:#1B3A6B;">VS</div>',
                unsafe_allow_html=True)
with pc2:
    st.markdown('<div class="filter-label">Player 2</div>', unsafe_allow_html=True)
    p2_opts   = [p for p in ALL_PLAYERS if p != sel_p1]
    default_p2= st.session_state.get("cmp_player2", p2_opts[0])
    default_p2= default_p2 if default_p2 in p2_opts else p2_opts[0]
    sel_p2    = st.selectbox("P2", p2_opts,
                             index=p2_opts.index(default_p2),
                             label_visibility="collapsed", key="sel_p2")

# Live role compatibility check
r1_raw = get_role(sel_p1)
r2_raw = get_role(sel_p2)
cmp_type, cmp_label = compare_mode_label(r1_raw, r2_raw)

if cmp_type is None:
    st.markdown(
        f'<div class="role-warn"><span class="role-warn-icon"></span>'
        f'<span class="role-warn-text">{cmp_label}</span></div>',
        unsafe_allow_html=True)

# Row 2 — Mode + filters
COMPARE_MODES = ["Overall","Recent Form","Country Wise",
                 "Opponent Wise","SENA Countries","Asian Countries"]
default_mode  = st.session_state.get("cmp_mode", "Overall")

mc1, mc2, mc3, mc4 = st.columns([4, 2, 2, 2])
_df1 = stats_dict.get(sel_p1, pd.DataFrame())
_df2 = stats_dict.get(sel_p2, pd.DataFrame())

_opps = sorted(set(
    list(_df1["Opponent"].dropna().unique() if not _df1.empty and "Opponent" in _df1 else []) +
    list(_df2["Opponent"].dropna().unique() if not _df2.empty and "Opponent" in _df2 else [])
))
_yrs = sorted(set(
    list(_df1["Match_date"].dt.year.dropna().astype(int).unique() if not _df1.empty else []) +
    list(_df2["Match_date"].dt.year.dropna().astype(int).unique() if not _df2.empty else [])
))

with mc1:
    st.markdown('<div class="filter-label">Compare Mode</div>', unsafe_allow_html=True)
    mode_idx = COMPARE_MODES.index(default_mode) if default_mode in COMPARE_MODES else 0
    sel_mode = st.selectbox("Mode", COMPARE_MODES, index=mode_idx,
                            label_visibility="collapsed")
with mc2:
    st.markdown('<div class="filter-label">Opponent</div>', unsafe_allow_html=True)
    sel_opp = st.selectbox("Opp", ["All"] + _opps, label_visibility="collapsed",
                           disabled=sel_mode not in ("Opponent Wise","Overall","Recent Form"))
with mc3:
    st.markdown('<div class="filter-label">Country</div>', unsafe_allow_html=True)
    if sel_mode == "SENA Countries":
        sel_ctr = st.selectbox("Ctr",["All SENA"]+sorted(SENA_COUNTRIES),
                               label_visibility="collapsed")
    elif sel_mode == "Asian Countries":
        sel_ctr = st.selectbox("Ctr",["All Asia"]+sorted(ASIA_COUNTRIES),
                               label_visibility="collapsed")
    else:
        _ctrs = sorted(set(
            list(_df1["Country"].dropna().unique() if not _df1.empty and "Country" in _df1 else []) +
            list(_df2["Country"].dropna().unique() if not _df2.empty and "Country" in _df2 else [])
        ))
        sel_ctr = st.selectbox("Ctr",["All"]+_ctrs, label_visibility="collapsed")

with mc4:
    st.markdown('<div class="filter-label">Year</div>', unsafe_allow_html=True)
    sel_year = st.selectbox("Year",["All"]+[str(y) for y in _yrs],
                            label_visibility="collapsed")

st.markdown('<div class="filter-label">&nbsp;</div>', unsafe_allow_html=True)
compare_btn = st.button("COMPARE", use_container_width=True)

# ── Session state gate ─────────────────────────────────────────
if compare_btn and cmp_type is not None:
    st.session_state.update(
        cmp_ok=True, cmp_p1=sel_p1, cmp_p2=sel_p2,
        cmp_mode_active=sel_mode,
        cmp_opp=sel_opp, cmp_ctr=sel_ctr, cmp_year=sel_year)

if not st.session_state.get("cmp_ok"):
    st.markdown(
        '<div class="no-data"><span class="icon"></span>'
        'Select two players and click <b>Compare</b></div>',
        unsafe_allow_html=True)
    st.stop()

# Restore from session
sel_p1   = st.session_state["cmp_p1"]
sel_p2   = st.session_state["cmp_p2"]
sel_mode = st.session_state["cmp_mode_active"]
sel_opp  = st.session_state["cmp_opp"]
sel_ctr  = st.session_state["cmp_ctr"]
sel_year = st.session_state["cmp_year"]

m1 = get_meta(sel_p1)
m2 = get_meta(sel_p2)
cmp_type, cmp_label = compare_mode_label(m1["role"], m2["role"])

if cmp_type is None:
    st.error(cmp_label); st.stop()

# Build filtered data
df1, bat1, bowl1 = build_df(sel_p1, sel_mode, sel_opp, sel_ctr, sel_year)
df2, bat2, bowl2 = build_df(sel_p2, sel_mode, sel_opp, sel_ctr, sel_year)

year_specific = (
    sel_year != "All"
    or sel_mode == "Recent Form"
    or sel_mode in ("SENA Countries", "Asian Countries")
    or (sel_mode == "Opponent Wise" and sel_opp not in ("All", ""))
    or (sel_mode == "Country Wise"  and sel_ctr not in ("All", ""))
    or (sel_mode == "Overall"       and sel_opp not in ("All", ""))
)

# Human-readable label used in chart titles when year_specific=True
if sel_mode == "Recent Form":
    filter_label = "Last 5 Matches"
elif sel_mode in ("SENA Countries", "Asian Countries"):
    grp = sel_ctr if sel_ctr not in ("All SENA", "All Asia", "All", "") else sel_mode
    filter_label = grp
elif sel_mode == "Opponent Wise" and sel_opp not in ("All", ""):
    filter_label = f"vs {sel_opp}"
elif sel_mode == "Country Wise" and sel_ctr not in ("All", ""):
    filter_label = sel_ctr
elif sel_year != "All":
    filter_label = str(sel_year)
else:
    filter_label = sel_mode

# ══════════════════════════════════════════════════════════════
#  SECTION 1 — MODE BANNER + PROFILES
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)

MODE_DESC = {
    "Overall":         "All available match data",
    "Recent Form":     "Last 5 matches",
    "Country Wise":    "Filtered by country",
    "Opponent Wise":   "vs specific opponent",
    "SENA Countries":  "SA · ENG · NZ · AUS",
    "Asian Countries": "IND · SL · BAN · PAK · UAE",
}
year_lbl = f" · {sel_year}" if sel_year != "All" else ""
opp_lbl  = f" vs {sel_opp}" if sel_opp != "All" else ""
ctr_lbl  = (f" · {sel_ctr}" if sel_ctr not in ("All","All SENA","All Asia","") else "")

st.markdown(
    f'<div class="mode-banner">'
    f'<span class="mode-pill">{sel_mode}</span>'
    f'<div>'
    f'<div class="mode-title">{sel_p1} vs {sel_p2}</div>'
    f'<div class="mode-sub">{cmp_label} &nbsp;·&nbsp; '
    f'{MODE_DESC.get(sel_mode,"")}'
    f'{opp_lbl}{ctr_lbl}{year_lbl}</div>'
    f'</div></div>',
    unsafe_allow_html=True)

# Profile cards
def profile_html(meta, name, player_num):
    cls  = "p1-card" if player_num == 1 else "p2-card"
    clr  = P1_CLR   if player_num == 1 else P2_CLR
    bcl  = badge_cls(meta["role"])
    lbl  = ('<span class="p1-badge">Player 1</span>' if player_num == 1
            else '<span class="p2-badge">Player 2</span>')
    img_src  = get_img_b64(meta["img"]) if meta["img"] else None
    img_html = (f'<img src="{img_src}" style="width:90px;height:90px;border-radius:50%;'
                f'object-fit:cover;border:3px solid {clr};display:block;margin:0 auto;">'
                if img_src else initials_avatar(name, clr))
    return f"""
    <div class="profile-card {cls}">
      {img_html}
      <div class="player-name">{name}</div>
      <div class="player-role-badge {bcl}">{meta['role']}</div>
      <div class="info-row"><span class="info-lbl">Batting</span>
        <span class="info-val">{meta['bat']}</span></div>
      <div class="info-row"><span class="info-lbl">Bowling</span>
        <span class="info-val">{meta['bowl']}</span></div>
    <div class="info-row"><span class="info-lbl">Date Of Birth</span>
        <span class="info-val"> {meta['dob']}</span></div>
      <div class="info-row"><span class="info-lbl">Debut Year</span>
        <span class="info-val">{meta['debut']}</span></div>
    </div>"""

p_l, vs_c, p_r = st.columns([5, 1.2, 5], gap="medium")
with p_l:
    st.markdown(profile_html(m1, sel_p1, 1), unsafe_allow_html=True)
with vs_c:
    n1 = role_norm(m1["role"]); n2 = role_norm(m2["role"])
    sub_vs = (
        "Bat vs Bat"    if cmp_type == "batting" and n1 != "all-rounder" and n2 != "all-rounder" else
        "Bowl vs Bowl"  if cmp_type == "bowling" and n1 != "all-rounder" and n2 != "all-rounder" else
        "Bat Stats"     if cmp_type == "batting"  else
        "Bowl Stats"    if cmp_type == "bowling"  else
        "All-Round"
    )
    st.markdown(
        f'<div class="vs-badge">'
        f'<div class="vs-circle">VS</div>'
        f'<div class="vs-sub">{sub_vs}</div>'
        f'</div>',
        unsafe_allow_html=True)
with p_r:
    st.markdown(profile_html(m2, sel_p2, 2), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SECTION 2 — STATS TABLE
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown(
    f'<div class="sec-title"><span class="dot"></span>'
    f'Statistical Comparison — {cmp_label} ({sel_mode})</div>',
    unsafe_allow_html=True)

if cmp_type == "batting":
    s1 = compute_bat(bat1)
    s2 = compute_bat(bat2)
    st.markdown(
        build_table(s1, s2, sel_p1, sel_p2, "Batting Statistics"),
        unsafe_allow_html=True)

elif cmp_type == "bowling":
    s1 = compute_bowl(bowl1)
    s2 = compute_bowl(bowl2)
    st.markdown(
        build_table(s1, s2, sel_p1, sel_p2, "Bowling Statistics"),
        unsafe_allow_html=True)

else:   # both (AR vs AR)
    sb1 = compute_bat(bat1);   sb2 = compute_bat(bat2)
    sw1 = compute_bowl(bowl1); sw2 = compute_bowl(bowl2)

    # Batting table
    st.markdown(
        build_table(sb1, sb2, sel_p1, sel_p2, " Batting Statistics"),
        unsafe_allow_html=True)
    st.markdown('<div style="height:20px;"></div>', unsafe_allow_html=True)
    # Bowling table
    st.markdown(
        build_table(sw1, sw2, sel_p1, sel_p2, " Bowling Statistics"),
        unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SECTION 3 — CHARTS
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="sec-title"><span class="dot"></span>'
            'Head-to-Head Charts</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
#  BATTING CHARTS
# ─────────────────────────────────────────────────────────────
def batting_charts(b1, b2, label1, label2, year_specific, filter_label):
    if not year_specific:
        ca, cb, cc = st.columns(3, gap="large")

        # Chart 1 — Year-wise Runs Line
        with ca:
            def yr_runs(bdf):
                if bdf.empty: return pd.Series(dtype=float)
                return bdf.assign(Y=bdf["Match_date"].dt.year).groupby("Y")["Runs"].sum()
            r1y = yr_runs(b1); r2y = yr_runs(b2)
            yrs = sorted(set(list(r1y.index)+list(r2y.index)))
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=yrs,y=[r1y.get(y,0) for y in yrs],
                name=label1,mode="lines+markers",
                line=dict(color=P1_CLR,width=2.5),
                marker=dict(size=9,color=P1_CLR,line=dict(color="white",width=2))))
            fig.add_trace(go.Scatter(x=yrs,y=[r2y.get(y,0) for y in yrs],
                name=label2,mode="lines+markers",
                line=dict(color=P2_CLR,width=2.5,dash="dot"),
                marker=dict(size=9,color=P2_CLR,line=dict(color="white",width=2))))
            pchart(fig,320,"Year-Wise Runs Trend")
            fig.update_xaxes(title_text="Year",tickformat="d")
            fig.update_yaxes(title_text="Runs")
            pch(fig)
        
        # Chart 2 — Average by Year Grouped Bar
        with cb:
            def yr_avg(bdf):
                if bdf.empty: return pd.DataFrame(columns=["Y","Avg"])
                out=[]
                for yr,g in bdf.assign(Y=bdf["Match_date"].dt.year).groupby("Y"):
                    a = calc_avg(g["Runs"],g.get("Not_out",pd.Series(0,index=g.index)),len(g))
                    out.append({"Y":yr,"Avg":a})
                return pd.DataFrame(out)
            a1y=yr_avg(b1); a2y=yr_avg(b2)
            fig2=go.Figure()
            fig2.add_trace(go.Bar(x=a1y["Y"] if not a1y.empty else [],
                y=a1y["Avg"] if not a1y.empty else [],name=label1,
                marker_color=P1_CLR,
                text=a1y["Avg"].round(1) if not a1y.empty else [],
                textposition="outside",textfont=dict(size=10)))
            fig2.add_trace(go.Bar(x=a2y["Y"] if not a2y.empty else [],
                y=a2y["Avg"] if not a2y.empty else [],name=label2,
                marker_color=P2_CLR,
                text=a2y["Avg"].round(1) if not a2y.empty else [],
                textposition="outside",textfont=dict(size=10)))
            fig2.update_layout(barmode="group")
            pchart(fig2,320,"Batting Average by Year")
            fig2.update_xaxes(title_text="Year",tickformat="d")
            fig2.update_yaxes(title_text="Average")
            pch(fig2)

        # Chart 3 — Scoring Pattern Stacked Bar
        with cc:
            def scpat(bdf):
                if bdf.empty: return {}
                return {
                    "1s/2s/3s": int(bdf["1s_2s_3s"].sum()) if "1s_2s_3s" in bdf.columns else 0,
                    "4s Runs":  int(bdf["4s"].sum())*4,
                    "6s Runs":  int(bdf["6s"].sum())*6,
                    "Dot Balls":int(bdf["Dots"].sum()) if "Dots" in bdf.columns else 0,
                }
            sp1=scpat(b1); sp2=scpat(b2)
            cats=["1s/2s/3s","4s Runs","6s Runs","Dot Balls"]
            clrs=[NAVY,AMBER,STEEL,"#D5DDE5"]
            fig3=go.Figure()
            for cat,clr in zip(cats,clrs):
                fig3.add_trace(go.Bar(name=cat,x=[label1,label2],
                    y=[sp1.get(cat,0),sp2.get(cat,0)],marker_color=clr))
            fig3.update_layout(barmode="stack")
            pchart(fig3,320,"Scoring Pattern Comparison")
            fig3.update_yaxes(title_text="Balls/Runs")
            pch(fig3)

    else:
        ca, cb, cc = st.columns(3, gap="large")

        # Chart 1 — Runs vs SR Scatter
        with ca:
            fig=go.Figure()
            for bdf,name,clr in [(b1,label1,P1_CLR),(b2,label2,P2_CLR)]:
                if not bdf.empty:
                    fig.add_trace(go.Scatter(
                        x=bdf["Runs"],
                        y=bdf["Strike_Rate"].replace(0,float("nan")),
                        mode="markers",name=name,
                        marker=dict(size=10,color=clr,opacity=.85,
                            line=dict(color="white",width=1.5))))
            pchart(fig,320,f"Batting Efficiency — {filter_label}")
            fig.update_xaxes(title_text="Runs")
            fig.update_yaxes(title_text="Strike Rate")
            pch(fig)

        # Chart 2 — Match Contribution Trend
        with cb:
            fig2=go.Figure()
            for bdf,name,clr in [(b1,label1,P1_CLR),(b2,label2,P2_CLR)]:
                if not bdf.empty:
                    td=bdf.sort_values("Match_date")
                    fig2.add_trace(go.Scatter(
                        x=list(range(1,len(td)+1)),y=td["Runs"],
                        mode="lines+markers",name=name,
                        line=dict(color=clr,width=2.5),
                        marker=dict(size=8,color=clr,line=dict(color="white",width=1.5))))
            pchart(fig2,320,f"Runs Trend Match by Match — {filter_label}")
            fig2.update_xaxes(title_text="Match #")
            fig2.update_yaxes(title_text="Runs")
            pch(fig2)

        # Chart 3 — Scoring Pattern
        with cc:
            def scpat2(bdf):
                if bdf.empty: return {}
                return {
                    "1s/2s/3s": int(bdf["1s_2s_3s"].sum()) if "1s_2s_3s" in bdf.columns else 0,
                    "4s Runs":  int(bdf["4s"].sum())*4,
                    "6s Runs":  int(bdf["6s"].sum())*6,
                    "Dot Balls":int(bdf["Dots"].sum()) if "Dots" in bdf.columns else 0,
                }
            sp1=scpat2(b1); sp2=scpat2(b2)
            cats=["1s/2s/3s","4s Runs","6s Runs","Dot Balls"]
            clrs=[NAVY,AMBER,STEEL,"#D5DDE5"]
            fig3=go.Figure()
            for cat,clr in zip(cats,clrs):
                fig3.add_trace(go.Bar(name=cat,x=[label1,label2],
                    y=[sp1.get(cat,0),sp2.get(cat,0)],marker_color=clr))
            fig3.update_layout(barmode="stack")
            pchart(fig3,320,f"Scoring Pattern — {filter_label}")
            fig3.update_yaxes(title_text="Balls/Runs")
            pch(fig3)


# ─────────────────────────────────────────────────────────────
#  BOWLING CHARTS
# ─────────────────────────────────────────────────────────────
def bowling_charts(bw1, bw2, label1, label2, year_specific, filter_label):
    if not year_specific:
        ca, cb, cc = st.columns(3, gap="large")

        # Chart 1 — Year-wise Wickets Line
        with ca:
            
            def yr_wkts(bdf):
                if bdf.empty: return pd.Series(dtype=float)
                return bdf.assign(Y=bdf["Match_date"].dt.year).groupby("Y")["Wickets"].sum()
            w1y=yr_wkts(bw1); w2y=yr_wkts(bw2)
            yrs=sorted(set(list(w1y.index)+list(w2y.index)))
            fig=go.Figure()
            fig.add_trace(go.Scatter(x=yrs,y=[w1y.get(y,0) for y in yrs],
                name=label1,mode="lines+markers",
                line=dict(color=P1_CLR,width=2.5),
                marker=dict(size=9,color=P1_CLR,line=dict(color="white",width=2))))
            fig.add_trace(go.Scatter(x=yrs,y=[w2y.get(y,0) for y in yrs],
                name=label2,mode="lines+markers",
                line=dict(color=P2_CLR,width=2.5,dash="dot"),
                marker=dict(size=9,color=P2_CLR,line=dict(color="white",width=2))))
            pchart(fig,320,"Year-Wise Wickets Trend")
            fig.update_xaxes(title_text="Year",tickformat="d")
            fig.update_yaxes(title_text="Wickets")
            pch(fig)

        # Chart 2 — Economy by Year Grouped Bar
        with cb:
            def yr_eco(bdf):
                if bdf.empty: return pd.DataFrame(columns=["Y","Eco"])
                out=[]
                for yr,g in bdf.assign(Y=bdf["Match_date"].dt.year).groupby("Y"):
                    out.append({"Y":yr,
                        "Eco":round(g["Economy"].replace(0,float("nan")).mean(),2)})
                return pd.DataFrame(out)
            e1y=yr_eco(bw1); e2y=yr_eco(bw2)
            fig2=go.Figure()
            fig2.add_trace(go.Bar(x=e1y["Y"] if not e1y.empty else [],
                y=e1y["Eco"] if not e1y.empty else [],name=label1,
                marker_color=P1_CLR,
                text=e1y["Eco"].round(2) if not e1y.empty else [],
                textposition="outside",textfont=dict(size=10)))
            fig2.add_trace(go.Bar(x=e2y["Y"] if not e2y.empty else [],
                y=e2y["Eco"] if not e2y.empty else [],name=label2,
                marker_color=P2_CLR,
                text=e2y["Eco"].round(2) if not e2y.empty else [],
                textposition="outside",textfont=dict(size=10)))
            fig2.update_layout(barmode="group")
            pchart(fig2,320,"Economy Rate by Year")
            fig2.update_xaxes(title_text="Year",tickformat="d")
            fig2.update_yaxes(title_text="Economy")
            pch(fig2)

        # Chart 3 — Bowling Control (Dot%, Economy, SR)
        with cc:
            def ctrl(bdf):
                if bdf.empty: return {}
                tb  = int(bdf["Balls_Bowled"].sum()) if "Balls_Bowled" in bdf.columns else 1
                db  = int(gcol(bdf,"Dot_Balls").sum())
                eco = round(bdf["Economy"].replace(0,float("nan")).mean(),2)
                wk  = int(bdf["Wickets"].sum())
                return {"Dot Ball %":round(db/max(tb,1)*100,1),
                        "Economy":eco, "Bowl SR":round(tb/max(wk,1),1)}
            c1s=ctrl(bw1); c2s=ctrl(bw2)
            cats=["Dot Ball %","Economy","Bowl SR"]
            fig3=go.Figure()
            fig3.add_trace(go.Bar(name=label1,x=cats,
                y=[c1s.get(k,0) for k in cats],marker_color=P1_CLR,
                text=[c1s.get(k,0) for k in cats],textposition="outside",
                textfont=dict(size=10)))
            fig3.add_trace(go.Bar(name=label2,x=cats,
                y=[c2s.get(k,0) for k in cats],marker_color=P2_CLR,
                text=[c2s.get(k,0) for k in cats],textposition="outside",
                textfont=dict(size=10)))
            fig3.update_layout(barmode="group")
            pchart(fig3,320,"Bowling Control Comparison")
            pch(fig3)

    else:
        ca, cb, cc = st.columns(3, gap="large")

        # Chart 1 — Wickets vs Economy Scatter
        with ca:
            fig=go.Figure()
            for bdf,name,clr in [(bw1,label1,P1_CLR),(bw2,label2,P2_CLR)]:
                if not bdf.empty:
                    fig.add_trace(go.Scatter(
                        x=bdf["Economy"].replace(0,float("nan")),
                        y=bdf["Wickets"],mode="markers",name=name,
                        marker=dict(size=10,color=clr,opacity=.85,
                            line=dict(color="white",width=1.5))))
            pchart(fig,320,f"Bowling Efficiency — {filter_label}")
            fig.update_xaxes(title_text="Economy")
            fig.update_yaxes(title_text="Wickets")
            pch(fig)

        # Chart 2 — Wickets Trend Match by Match
        with cb:
            fig2=go.Figure()
            for bdf,name,clr in [(bw1,label1,P1_CLR),(bw2,label2,P2_CLR)]:
                if not bdf.empty:
                    td=bdf.sort_values("Match_date")
                    fig2.add_trace(go.Scatter(
                        x=list(range(1,len(td)+1)),y=td["Wickets"],
                        mode="lines+markers",name=name,
                        line=dict(color=clr,width=2.5),
                        marker=dict(size=8,color=clr,line=dict(color="white",width=1.5))))
            pchart(fig2,320,f"Wickets Trend — {filter_label}")
            fig2.update_xaxes(title_text="Match #")
            fig2.update_yaxes(title_text="Wickets")
            pch(fig2)

        # Chart 3 — Bowling Control
        with cc:
            def ctrl2(bdf):
                if bdf.empty: return {}
                tb  = int(bdf["Balls_Bowled"].sum()) if "Balls_Bowled" in bdf.columns else 1
                db  = int(gcol(bdf,"Dot_Balls").sum())
                eco = round(bdf["Economy"].replace(0,float("nan")).mean(),2)
                wk  = int(bdf["Wickets"].sum())
                return {"Dot Ball %":round(db/max(tb,1)*100,1),
                        "Economy":eco,"Bowl SR":round(tb/max(wk,1),1)}
            c1s=ctrl2(bw1); c2s=ctrl2(bw2)
            cats=["Dot Ball %","Economy","Bowl SR"]
            fig3=go.Figure()
            fig3.add_trace(go.Bar(name=label1,x=cats,
                y=[c1s.get(k,0) for k in cats],marker_color=P1_CLR,
                text=[c1s.get(k,0) for k in cats],textposition="outside",
                textfont=dict(size=10)))
            fig3.add_trace(go.Bar(name=label2,x=cats,
                y=[c2s.get(k,0) for k in cats],marker_color=P2_CLR,
                text=[c2s.get(k,0) for k in cats],textposition="outside",
                textfont=dict(size=10)))
            fig3.update_layout(barmode="group")
            pchart(fig3,320,f"Bowling Control — {filter_label}")
            pch(fig3)


# ─────────────────────────────────────────────────────────────
#  ALL-ROUNDER RADAR
# ─────────────────────────────────────────────────────────────
def ar_radar_chart(b1, bw1, b2, bw2, label1, label2):
    def scores(bdf, bowl_df):
        avg_ = calc_avg(bdf["Runs"] if not bdf.empty else pd.Series([0]),
                        bdf.get("Not_out",pd.Series(0,index=bdf.index)) if not bdf.empty else pd.Series([0]),
                        max(len(bdf),1))
        sr_  = round(bdf["Strike_Rate"].replace(0,float("nan")).mean(),1) if not bdf.empty else 0
        runs_= int(bdf["Runs"].sum()) if not bdf.empty else 0
        wk_  = int(bowl_df["Wickets"].sum()) if not bowl_df.empty else 0
        eco_ = round(bowl_df["Economy"].replace(0,float("nan")).mean(),2) if not bowl_df.empty else 0
        return {
            "Bat Average":   min(avg_/100*100, 100),
            "Strike Rate":   min(sr_/200*100, 100),
            "Runs":          min(runs_/500*100, 100),
            "Wickets":       min(wk_/20*100, 100),
            "Economy (inv)": max(0,(10-eco_)/10*100),
        }
    r1s=scores(b1,bw1); r2s=scores(b2,bw2)
    cats=list(r1s.keys())
    v1=[r1s[k] for k in cats]+[r1s[cats[0]]]
    v2=[r2s[k] for k in cats]+[r2s[cats[0]]]
    c_r=cats+[cats[0]]
    _, radar_col, _ = st.columns([1, 4, 1])
    with radar_col:
        fig=go.Figure()
        fig.add_trace(go.Scatterpolar(r=v1,theta=c_r,fill="toself",name=label1,
            line=dict(color=P1_CLR,width=2.5),fillcolor=P1_FILL,
            marker=dict(color=P1_CLR,size=8,line=dict(color="white",width=2))))
        fig.add_trace(go.Scatterpolar(r=v2,theta=c_r,fill="toself",name=label2,
            line=dict(color=P2_CLR,width=2.5,dash="dot"),fillcolor=P2_FILL,
            marker=dict(color=P2_CLR,size=8,line=dict(color="white",width=2))))
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",height=400,
            margin=dict(l=20,r=20,t=50,b=10),
            title=dict(text="All-Round Skill Comparison (Radar)",
                font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
            polar=dict(
                radialaxis=dict(range=[0,100],visible=True,
                    gridcolor="#f0f3f8",linecolor=BORDER,tickfont=dict(size=7)),
                angularaxis=dict(tickfont=dict(family="Inter",size=9,color=NAVY))),
            legend=dict(bgcolor="rgba(255,255,255,.9)",
                bordercolor=BORDER,font=dict(size=10)),
        )
        pch(fig)


# ─────────────────────────────────────────────────────────────
#  RENDER CHARTS by cmp_type
# ─────────────────────────────────────────────────────────────
if cmp_type == "batting":
    batting_charts(bat1, bat2, sel_p1, sel_p2, year_specific, filter_label)

elif cmp_type == "bowling":
    bowling_charts(bowl1, bowl2, sel_p1, sel_p2, year_specific, filter_label)

else:  # both (AR vs AR)
    st.markdown(
        '<div class="cmp-mode-pill">Batting Charts</div>',
        unsafe_allow_html=True)
    batting_charts(bat1, bat2, sel_p1, sel_p2, year_specific, filter_label)

    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown(
        '<div class="cmp-mode-pill">Bowling Charts</div>',
        unsafe_allow_html=True)
    bowling_charts(bowl1, bowl2, sel_p1, sel_p2, year_specific, filter_label)

    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    st.markdown(
        '<div class="cmp-mode-pill">Overall Skill Radar</div>',
        unsafe_allow_html=True)
    ar_radar_chart(bat1, bowl1, bat2, bowl2, sel_p1, sel_p2)

# ══════════════════════════════════════════════════════════════
#  ACTION BAR + FOOTER
# ══════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="action-bar">
  <a class="act-btn act-report" href="/generate_report" target="_self">
    &nbsp;Generate Comparison Report
  </a>
  <div style="color:#8899aa;font-size:.78rem;margin-top:10px;">
    {sel_p1} vs {sel_p2} &nbsp;·&nbsp;
    {sel_mode} &nbsp;·&nbsp;
    {sel_year if sel_year != "All" else "All Years"}
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="footer">
  Cricket Player Performance Analysis Dashboard &nbsp;|&nbsp; Built with Streamlit
</div>
""", unsafe_allow_html=True)