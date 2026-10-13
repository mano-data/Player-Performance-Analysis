import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import openpyxl, base64, os, math

# ── PAGE CONFIG ───────────────────────────────────────────────
st.set_page_config(
    page_title="Strength & Weakness | Cricket Analysis",
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

# ── BRAND PALETTE  (identical to all other pages) ────────────
NAVY  = "#1B3A6B"
AMBER = "#E8A020"
STEEL = "#4A6FA5"
GREEN = "#1E7A34"
RED   = "#C0392B"
BORDER= "#dde3ee"

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
.sec-title{font-family:'Poppins',sans-serif;font-size:1rem;font-weight:700;color:#1B3A6B;
    margin:0 0 14px 0;padding-bottom:8px;border-bottom:2px solid #e8edf5;
    display:flex;align-items:center;gap:8px;}
.sec-title .dot{width:8px;height:8px;background:#1B3A6B;border-radius:50%;display:inline-block;}

/* ─ PROFILE CARD ─ */
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

/* ─ INSIGHT CARDS ─ */
.ins-strength{background:#f0f8ee;border:1px solid #a8d4a0;
    border-left:4px solid #1E7A34;border-radius:8px;
    padding:11px 15px;display:flex;align-items:flex-start;gap:11px;margin-bottom:8px;}
.ins-weakness{background:#fde8e8;border:1px solid #f0b8b8;
    border-left:4px solid #C0392B;border-radius:8px;
    padding:11px 15px;display:flex;align-items:flex-start;gap:11px;margin-bottom:8px;}
.ins-neutral{background:#f7f9fc;border:1px solid #dde3ee;
    border-left:4px solid #E8A020;border-radius:8px;
    padding:11px 15px;display:flex;align-items:flex-start;gap:11px;margin-bottom:8px;}
.ins-icon{font-size:1.05rem;flex-shrink:0;margin-top:2px;}
.ins-text{font-size:.83rem;color:#2c3e50;line-height:1.5;}
.ins-text strong{color:#1B3A6B;}
.ins-section-lbl{font-family:'Poppins',sans-serif;font-size:.72rem;font-weight:700;
    letter-spacing:.8px;text-transform:uppercase;margin:12px 0 7px 0;}
.str-lbl{color:#1E7A34;}
.wk-lbl{color:#C0392B;}
.met-lbl{color:#6B7A8D;}

/* ─ KPI CARDS ─ */
.kpi-card{background:#fff;border:1px solid #dde3ee;border-radius:10px;
    padding:18px 14px 14px 14px;text-align:center;
    box-shadow:0 2px 8px rgba(27,58,107,.05);
    transition:box-shadow .2s,transform .2s;position:relative;overflow:hidden;}
.kpi-card:hover{box-shadow:0 6px 18px rgba(27,58,107,.12);transform:translateY(-2px);}
.kpi-card::after{content:'';position:absolute;bottom:0;left:0;right:0;height:3px;
    background:#1B3A6B;}
.kpi-card.hl::after{background:#E8A020;}
.kpi-card.hl-g::after{background:#1E7A34;}
.kpi-val{font-family:'Poppins',sans-serif;font-size:1.65rem;font-weight:700;
    color:#1B3A6B;line-height:1;margin-bottom:6px;}
.kpi-lbl{font-size:.7rem;color:#8899aa;font-weight:600;
    letter-spacing:.6px;text-transform:uppercase;}
.kpi-sub{font-size:.68rem;color:#aab0bc;margin-top:3px;}

/* ─ CHART CARD ─ */
.chart-card{background:#fff;border:1px solid #dde3ee;border-radius:12px;
    padding:20px 18px 10px 18px;
    box-shadow:0 2px 10px rgba(27,58,107,.06);margin-bottom:20px;}

/* ─ DIVIDER / NO DATA ─ */
.divider{border:none;height:1px;background:#e8edf5;margin:28px 0;}
.no-data{text-align:center;padding:50px;color:#8899aa;font-size:.9rem;}
.no-data .icon{font-size:2.5rem;display:block;margin-bottom:10px;}

/* ─ ACTION BAR / FOOTER ─ */
.action-bar{background:#f7f9fc;border-top:1px solid #dde3ee;
    padding:26px 60px;margin:32px -60px 0 -60px;
    display:flex;align-items:center;justify-content:space-between;}
.act-btn{display:inline-flex;align-items:center;gap:8px;padding:11px 28px;
    border-radius:7px;font-family:'Poppins',sans-serif;font-weight:700;
    font-size:.88rem;letter-spacing:.5px;
    cursor:pointer;text-decoration:none !important;border:none;}
.act-report{background:#fff;color:#1B3A6B !important;border:2px solid #1B3A6B;}
.act-report:hover{background:#EBF3FF;}
.act-compare{background:#1B3A6B;color:#fff !important;}
.act-compare:hover{background:#243f7a;}
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

# ── UTILITY ───────────────────────────────────────────────────
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
    return round(tot/dis, 2) if dis > 0 else round(tot, 2)

def pchart(fig, h=340, title=""):
    """Consistent chart theme used across all pages."""
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#fff",
        font=dict(family="Inter", color="#2c3e50", size=11),
        height=h, margin=dict(l=10,r=10,t=44,b=10),
        title=dict(text=title,
            font=dict(family="Poppins",size=13,color=NAVY),
            x=0.02, xanchor="left"),
        legend=dict(bgcolor="rgba(255,255,255,.9)",
            bordercolor=BORDER, borderwidth=1, font=dict(size=10)),
    )
    fig.update_xaxes(gridcolor="#f0f3f8",linecolor=BORDER,
        tickfont=dict(size=9),zeroline=False)
    fig.update_yaxes(gridcolor="#f0f3f8",linecolor=BORDER,
        tickfont=dict(size=9),zeroline=False)
    return fig

def kcard(val, lbl, sub="", v=""):
    return (f'<div class="kpi-card {v}"><div class="kpi-val">{val}</div>'
            f'<div class="kpi-lbl">{lbl}</div>'
            + (f'<div class="kpi-sub">{sub}</div>' if sub else "") + '</div>')

def ins_s(icon, txt):
    return (f'<div class="ins-strength"><span class="ins-icon">{icon}</span>'
            f'<span class="ins-text">{txt}</span></div>')

def ins_w(icon, txt):
    return (f'<div class="ins-weakness"><span class="ins-icon">{icon}</span>'
            f'<span class="ins-text">{txt}</span></div>')

def ins_n(icon, txt):
    return (f'<div class="ins-neutral"><span class="ins-icon">{icon}</span>'
            f'<span class="ins-text">{txt}</span></div>')

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

# ══════════════════════════════════════════════════════════════
#  WAGON WHEEL  
# ══════════════════════════════════════════════════════════════
def make_wagon_wheel(bat_df_in, player_name="", h=420):
    """
    Pure-Plotly 8-sector wagon wheel matching CricViz/ESPNcricinfo aesthetics.
    No pitch, no image. Uses the project's navy palette.
    """
    # ── Zone layout (batsman faces UP, off-side = LEFT, leg-side = RIGHT)
    # Angles: standard math convention (0=right, 90=top, CCW positive)
    # Top-left quadrant  = off-side  | top-right quadrant  = leg-side
    # Bot-left quadrant  = off-side  | bot-right quadrant  = leg-side
    #
    #          Long Off (112.5°) | Long On (67.5°)
    #   Cover/MidOff (157.5°)   | MidWkt (22.5°)
    #         Point  (202.5°)   | SqLeg  (337.5°)
    #       ThirdMan (247.5°)   | FineLeg(292.5°)
    #
    ZONES = [
        #  zone_label          data_col          centre_angle_deg
        ("Long Off",        "Runs_Longoff",         112.5),
        ("Cover",           "Runs_Covers",          157.5),
        ("Point",           "Runs_Point",           202.5),
        ("Third Man",       "Runs_Thirdman",        247.5),
        ("Fine Leg",        "Runs_Fineleg",         292.5),
        ("Square Leg",      "Runs_Square",          337.5),
        ("Mid Wicket",      "Runs_Midwct",           22.5),
        ("Long On",         "Runs_Longon",           67.5),
    ]

    N          = len(ZONES)
    SECTOR_DEG = 360.0 / N          # 45° each
    ARC_PTS    = 80                  # arc smoothness
    R          = 1.0                 # wheel radius
    CX, CY     = 0.0, 0.0           # centre

    # ── aggregate per-zone run totals ─────────────────────────
    vals = []
    for _, col, _ in ZONES:
        vals.append(int(bat_df_in[col].sum()) if col in bat_df_in.columns else 0)

    total_runs = sum(vals)
    max_v      = max(vals) if max(vals) > 0 else 1

    # ── colour: light blue (low) → deep navy (high) ───────────
    # Light: rgb(214, 232, 247)  #D6E8F7
    # Dark:  rgb( 27,  58, 107)  #1B3A6B
    def fill_color(v):
        t = v / max_v
        r = int(214 + ( 27 - 214) * t)
        g = int(232 + ( 58 - 232) * t)
        b = int(247 + (107 - 247) * t)
        return f"rgb({r},{g},{b})"

    def num_color(v):
        """White on dark, navy on light — threshold 55%."""
        return "#ffffff" if v / max_v >= 0.55 else "#1B3A6B"

    def rim_color(v):
        """Zone label colour."""
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

        # ① Filled sector
        fig.add_trace(go.Scatter(
            x=poly_x, y=poly_y,
            fill="toself",
            fillcolor=fill_color(v),
            line=dict(color="white", width=2.8),
            mode="lines",
            hovertemplate=f"<b>{zone_lbl}</b><br>{v} runs<extra></extra>",
            showlegend=False,
        ))

        # ② Large run-value number — at 62% radius
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

        # ③ Zone label — at 84% radius (near outer edge)
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

    # ── OFF-SIDE / LEG-SIDE labels (horizontal midline) ───────
    # Placed inside the wheel at the 0° midline (y = 0)
    for lx_, txt_ in [(-0.46, "OFF-SIDE"), (0.46, "LEG-SIDE")]:
        fig.add_annotation(
            x=lx_, y=0.0,
            text=f"<b>{txt_}</b>",
            showarrow=False,
            font=dict(family="Inter", size=9,
                      color="rgba(255,255,255,0.75)"),
            xanchor="center", yanchor="middle",
        )

    # ── Colour-scale legend strip (below wheel) ───────────────
    STRIP_Y0, STRIP_Y1 = -1.18, -1.08
    for s in range(30):
        t  = s / 29
        r_ = int(214 + (27  - 214) * t)
        g_ = int(232 + (58  - 232) * t)
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

    # ── Subtitle (player name + total) ───────────────────────
    fig.add_annotation(
        x=0, y=1.20,
        text=(f"<span style='font-size:10.5px;color:#6B7A8D'>"
              f"{player_name}&nbsp;&nbsp;•&nbsp;&nbsp;"
              f"Total runs: <b>{total_runs}</b></span>"),
        showarrow=False,
        font=dict(family="Inter", size=11, color="#6B7A8D"),
        xanchor="center",
    )

    # ── Layout ────────────────────────────────────────────────
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        height=h,
        margin=dict(l=10, r=10, t=44, b=48),
        title=dict(
            text="Scoring Zone Strength — Wagon Wheel",
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


# ══════════════════════════════════════════════════════════════
#  NAVBAR
# ══════════════════════════════════════════════════════════════
st.markdown("""
<div class="top-header">
  <div class="header-brand"> Cricket Player<span> Performance Analysis</span></div>
  <div class="header-nav">
    <a class="h-btn" href="/" target="_self"> Back</a>
    <a class="h-btn" href="/userhome" target="_self"> Home</a>
    <a class="h-btn" href="/admin" target="_self"> Admin</a>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="page-title-bar">
  <h2> Strength &amp; Weakness Analysis</h2>
  <p>Deep-dive scoring zones, dismissal patterns, bowling pressure index and performance splits</p>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  FILTER PANEL
# ══════════════════════════════════════════════════════════════
fc1, fc2, fc3 = st.columns([3, 3, 2])

with fc1:
    st.markdown('<div class="filter-label"> Player</div>', unsafe_allow_html=True)
    sel_player = st.selectbox("Player", sorted(stats_dict.keys()),
                              label_visibility="collapsed")

with fc2:
    st.markdown('<div class="filter-label"> Country</div>', unsafe_allow_html=True)
    _raw  = stats_dict.get(sel_player, pd.DataFrame())
    _ctrs = (sorted([c for c in _raw["Country"].dropna().unique() if str(c).strip()])
             if not _raw.empty and "Country" in _raw else [])
    sel_ctr = st.selectbox("Country", ["All"] + _ctrs, label_visibility="collapsed")

with fc3:
    st.markdown('<div class="filter-label"> Year</div>', unsafe_allow_html=True)
    _yrs = (["All"] + sorted([int(y) for y in
             _raw["Match_date"].dropna().dt.year.unique()])
            if not _raw.empty else ["All"])
    sel_year = st.selectbox("Year", _yrs, label_visibility="collapsed")


st.markdown('<div class="filter-label">&nbsp;</div>', unsafe_allow_html=True)
analyze = st.button(" ANALYZE NOW", use_container_width=True)



# ── SESSION STATE ─────────────────────────────────────────────
if analyze:
    st.session_state.update(sw_ok=True, sw_player=sel_player,
                            sw_ctr=sel_ctr, sw_year=sel_year)

if not st.session_state.get("sw_ok"):
    st.markdown('<div class="no-data"><span class="icon"></span>'
                'Select a player &amp; filters, then click '
                '<b>Analyze Now</b></div>', unsafe_allow_html=True)
    st.stop()

sel_player = st.session_state["sw_player"]
sel_ctr    = st.session_state["sw_ctr"]
sel_year   = st.session_state["sw_year"]
df_raw     = stats_dict.get(sel_player, pd.DataFrame())

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
bcls   = ("badge-bowler" if is_bowl else
          ("badge-allrounder" if is_all else "badge-batsman"))

# ── FILTER WORKING DATAFRAME ──────────────────────────────────
if df_raw.empty:
    st.markdown('<div class="no-data"><span class="icon"></span>'
                'No data for this player.</div>', unsafe_allow_html=True)
    st.stop()

wdf = df_raw.copy()
if sel_ctr  != "All": wdf = wdf[wdf["Country"]  == sel_ctr]
if sel_year != "All": wdf = wdf[wdf["Match_date"].dt.year == int(sel_year)]
wdf = wdf.sort_values("Match_date")

bat_df  = bat_rows(wdf)
bowl_df = bowl_rows(wdf)

if bat_df.empty and bowl_df.empty:
    st.markdown('<div class="no-data"><span class="icon"></span>'
                'No data for selected filters.</div>',
                unsafe_allow_html=True)
    st.stop()

total_m   = len(wdf)
scope_lbl = ((sel_ctr if sel_ctr != "All" else "All Countries") +
             (f" · {sel_year}" if sel_year != "All" else ""))

# ══════════════════════════════════════════════════════════════
#  SECTION 1 — PROFILE + INSIGHTS
# ══════════════════════════════════════════════════════════════
col_p, col_i = st.columns([1, 2.2], gap="large")

# ── Left: Profile Card ────────────────────────────────────────
with col_p:
    st.markdown('<div class="sec-title"><span class="dot"></span>'
                'Player Profile</div>', unsafe_allow_html=True)
    img_src  = get_img_b64(p_img) if p_img else None
    img_html = (f'<img src="{img_src}" style="width:100px;height:100px;'
                f'border-radius:50%;object-fit:cover;border:3px solid #1B3A6B;'
                f'display:block;margin:0 auto;">'
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
        <span class="info-val">{scope_lbl}</span></div>
    </div>""", unsafe_allow_html=True)

# ── Right: Insights ───────────────────────────────────────────
with col_i:
    st.markdown('<div class="sec-title"><span class="dot"></span>'
                'Strength &amp; Weakness Insights</div>', unsafe_allow_html=True)
    ih = ""  # insight HTML accumulator

    # ════════════════ BATSMAN ════════════════
    if is_bat and not is_all:
        ih += '<div class="ins-section-lbl str-lbl"> STRENGTHS</div>'

        # 1. Best scoring zones
        ZONE_MAP = {
            "Long On":    "Runs_Longon",   "Long Off":   "Runs_Longoff",
            "Cover":      "Runs_Covers",   "Point":      "Runs_Point",
            "Third Man":  "Runs_Thirdman", "Fine Leg":   "Runs_Fineleg",
            "Square Leg": "Runs_Square",   "Mid Wicket": "Runs_Midwct",
        }
        zv = {k: int(bat_df[v].sum()) for k, v in ZONE_MAP.items()
              if v in bat_df.columns}
        if zv:
            top2 = sorted(zv, key=zv.get, reverse=True)[:2]
            ih += ins_s("",
                f"Strongest scoring zones: <strong>{top2[0]}</strong> "
                f"({zv[top2[0]]} runs) &amp; <strong>{top2[1]}</strong> "
                f"({zv[top2[1]]} runs)")

        # 2. Boundary dependency
        tot_r  = int(bat_df["Runs"].sum())
        bnd_r  = int(bat_df["4s"].sum()*4 + bat_df["6s"].sum()*6)
        bnd_pct= round(bnd_r / max(tot_r, 1) * 100, 1)
        ih += ins_s("",
            f"Boundary dependency: <strong>{bnd_pct}%</strong> of runs "
            f"from boundaries ({bnd_r} of {tot_r} total runs)")

        # 3. Strike rotation
        singles = (int(bat_df["1s_2s_3s"].sum())
                   if "1s_2s_3s" in bat_df.columns else 0)
        bf_tot  = int(bat_df["Balls_faced"].sum()) if "Balls_faced" in bat_df.columns else 1
        sr_pct  = round(singles / max(tot_r, 1) * 100, 1)
        ih += (ins_s if sr_pct >= 30 else ins_n)("",
            f"Strike rotation: <strong>{sr_pct}%</strong> of runs via "
            f"singles/twos/threes "
            f"({'good ' if sr_pct>=30 else 'needs improvement '})")

        ih += '<div class="ins-section-lbl wk-lbl"> WEAKNESSES</div>'

        # 4. Most dismissals by length
        dl = bat_df[~bat_df["Dismissed_Length"].astype(str).isin(
            ["Not Applicable","0","nan",""])]
        if not dl.empty:
            dlg  = dl["Dismissed_Length"].value_counts()
            weak = dlg.index[0]
            ih  += ins_w("",
                f"Most vulnerable to <strong>{weak}</strong> deliveries "
                f"({dlg.iloc[0]} dismissals) — key weakness to address")

        # 5. Most dismissed by bowler
        da = bat_df[~bat_df["Dismissed_against"].astype(str).isin(
            ["Not Applicable","0","nan",""])]
        if not da.empty:
            dag = da["Dismissed_against"].value_counts()
            ih += ins_w("",
                f"Most dismissed by: <strong>{dag.index[0]}</strong> "
                f"({dag.iloc[0]} times)")

        ih += '<div class="ins-section-lbl met-lbl"> METRICS</div>'

        # 6. Dot ball pressure
        dots   = int(bat_df["Dots"].sum()) if "Dots" in bat_df.columns else 0
        dot_pct= round(dots / max(bf_tot, 1) * 100, 1)
        ih += (ins_w if dot_pct > 45 else ins_n)("",
            f"Dot ball pressure: <strong>{dot_pct}%</strong> of balls "
            f"are dots ({dots} / {bf_tot}) — "
            f"{'high pressure ' if dot_pct>45 else 'under control '}")

        # 7. Innings split
        i1 = bat_df[bat_df["Innings_no"]==1]
        i2 = bat_df[bat_df["Innings_no"]==2]
        if not i1.empty and not i2.empty:
            a1 = calc_avg(i1["Runs"],
                i1.get("Not_out",pd.Series(0,index=i1.index)), len(i1))
            a2 = calc_avg(i2["Runs"],
                i2.get("Not_out",pd.Series(0,index=i2.index)), len(i2))
            better = "1st innings" if a1 >= a2 else "2nd innings"
            ih += ins_n("",
                f"1st Inn avg <strong>{a1}</strong> vs "
                f"2nd Inn avg <strong>{a2}</strong> — "
                f"performs better in <strong>{better}</strong>")

    # ════════════════ BOWLER ════════════════
    elif is_bowl and not is_all:
        ih += '<div class="ins-section-lbl str-lbl"> STRENGTHS</div>'

        t_balls = (int(bowl_df["Balls_Bowled"].sum())
                   if "Balls_Bowled" in bowl_df else 0)
        dots_b  = int(get_col(bowl_df,"Dot_Balls").sum())
        pi_v    = round(dots_b / max(t_balls, 1) * 100, 1)
        ih += ins_s("",
            f"Pressure Index: <strong>{pi_v}%</strong> dot balls "
            f"({dots_b}/{t_balls}) — "
            f"{'elite ' if pi_v>=40 else 'building '}")

        fd  = int(bowl_df["Full_Dismissals"].sum()) if "Full_Dismissals" in bowl_df else 0
        gd  = int(get_col(bowl_df,"Good_Dissmisals","Good_Dismissals").sum())
        sd  = int(get_col(bowl_df,"Short_Dissmisals","Short_Dismissals").sum())
        bv  = max(fd, gd, sd)
        bn  = {fd:"Full Length", gd:"Good Length", sd:"Short Pitch"}.get(bv,"Good Length")
        ih += ins_s("",
            f"Strongest length: <strong>{bn}</strong> "
            f"({bv} dismissals)")

        rhw = int(bowl_df["Wicket_RHB"].sum()) if "Wicket_RHB" in bowl_df else 0
        lhw = int(bowl_df["Wicket_LHB"].sum()) if "Wicket_LHB" in bowl_df else 0
        dom = "Right-Handed" if rhw >= lhw else "Left-Handed"
        ih += ins_s("",
            f"Stronger vs <strong>{dom}</strong> batters "
            f"(RHB {rhw} wkts | LHB {lhw} wkts)")

        ih += '<div class="ins-section-lbl wk-lbl"> WEAKNESSES</div>'

        we = bowl_df[bowl_df["Match_result"]=="Won"]["Economy"].replace(0,float("nan")).mean()
        le = bowl_df[bowl_df["Match_result"]=="Lose"]["Economy"].replace(0,float("nan")).mean()
        if pd.notnull(we) and pd.notnull(le):
            if le > we:
                ih += ins_w("",
                    f"Economy rises in losses: <strong>{le:.2f}</strong> "
                    f"vs wins <strong>{we:.2f}</strong>")
            else:
                ih += ins_s("",
                    f"Tighter economy even in losses "
                    f"({le:.2f}) vs wins ({we:.2f}) ")

        ext = int(get_col(bowl_df,"Extras").sum())
        ih += (ins_w if ext > 20 else ins_n)("",
            f"Extras conceded: <strong>{ext}</strong> — "
            f"{'concern ' if ext>20 else 'well controlled '}")

        ih += '<div class="ins-section-lbl met-lbl">METRICS</div>'
        tw  = int(bowl_df["Wickets"].sum())
        eco = round(bowl_df["Economy"].replace(0,float("nan")).mean(), 2)
        bsr = round(t_balls / max(tw, 1), 1)
        ih += ins_n("",
            f"Summary: <strong>{tw}</strong> wickets | "
            f"Eco <strong>{eco}</strong> | SR <strong>{bsr}</strong>")

    # ════════════════ ALL-ROUNDER ════════════════
    else:
        ih += '<div class="ins-section-lbl str-lbl"> STRENGTHS</div>'
        ZONE_MAP2 = {
            "Long On":"Runs_Longon","Long Off":"Runs_Longoff",
            "Cover":"Runs_Covers","Point":"Runs_Point",
            "Third Man":"Runs_Thirdman","Fine Leg":"Runs_Fineleg",
            "Square Leg":"Runs_Square","Mid Wicket":"Runs_Midwct",
        }
        zv2 = {k: int(bat_df[v].sum()) for k,v in ZONE_MAP2.items()
               if not bat_df.empty and v in bat_df.columns}
        if zv2:
            t2 = sorted(zv2, key=zv2.get, reverse=True)[:2]
            ih += ins_s("",
                f"Top batting zones: <strong>{t2[0]}</strong> "
                f"({zv2[t2[0]]}) &amp; <strong>{t2[1]}</strong> ({zv2[t2[1]]})")
        if not bowl_df.empty:
            tb2 = int(bowl_df["Balls_Bowled"].sum()) if "Balls_Bowled" in bowl_df else 0
            db2 = int(get_col(bowl_df,"Dot_Balls").sum())
            pi2 = round(db2/max(tb2,1)*100,1)
            ih += ins_s("",
                f"Bowling pressure index: <strong>{pi2}%</strong> dot balls")

        ih += '<div class="ins-section-lbl wk-lbl"> WEAKNESSES</div>'
        if not bat_df.empty:
            da2 = bat_df[~bat_df["Dismissed_against"].astype(str).isin(
                ["Not Applicable","0","nan",""])]
            if not da2.empty:
                dg2 = da2["Dismissed_against"].value_counts()
                ih += ins_w("",
                    f"Most dismissed by <strong>{dg2.index[0]}</strong> "
                    f"({dg2.iloc[0]} times)")
        if not bowl_df.empty:
            rr = int(bowl_df["Runs_RHB"].sum()) if "Runs_RHB" in bowl_df else 0
            lr = int(bowl_df["Runs_LHB"].sum()) if "Runs_LHB" in bowl_df else 0
            wk = "Right-Handed" if rr > lr else "Left-Handed"
            ih += ins_w("",
                f"More runs conceded vs <strong>{wk}</strong> batters "
                f"(RHB {rr} | LHB {lr} runs)")

        ih += '<div class="ins-section-lbl met-lbl"> METRICS</div>'
        if not bat_df.empty and not bowl_df.empty:
            tr3 = int(bat_df["Runs"].sum())
            tw3 = int(bowl_df["Wickets"].sum())
            ih += ins_n("",
                f"All-round: <strong>{tr3}</strong> runs + "
                f"<strong>{tw3}</strong> wickets in {total_m} matches")

    st.markdown(ih or '<div class="no-data" style="padding:20px">'
                'No insights available.</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SECTION 2 — KPI CARDS
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="sec-title"><span class="dot"></span>'
            'Key Performance Indicators</div>', unsafe_allow_html=True)

if is_bat and not is_all:
    inn  = len(bat_df)
    tr   = int(bat_df["Runs"].sum())
    no_s = bat_df.get("Not_out", pd.Series(0,index=bat_df.index))
    avg  = calc_avg(bat_df["Runs"], no_s, inn)
    avsr = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(), 1)
    f4   = int(bat_df["4s"].sum()); f6 = int(bat_df["6s"].sum())
    fif  = int(bat_df["50s"].sum()) if "50s" in bat_df.columns else 0
    hun  = int(bat_df["100s"].sum()) if "100s" in bat_df.columns else 0
    hs   = (f"{int(bat_df.loc[bat_df['Runs'].idxmax(),'Runs'])}"
            f"{'*' if int(bat_df.loc[bat_df['Runs'].idxmax(),'Not_out']) else ''}"
            if not bat_df.empty else "0")
    kpis = [
        (str(inn),       "Innings",    scope_lbl, ""),
        (str(tr),        "Total Runs", "Aggregate","hl"),
        (str(avg),       "Average",    "Runs/dis", ""),
        (str(avsr),      "Strike Rate","Avg SR",   ""),
        (hs,             "Best Score", "High inn", "hl"),
        (f"{f4}/{f6}",   "4s / 6s",   "Boundaries",""),
        (f"{fif}/{hun}", "50s/100s",  "Milestones","hl-g"),
    ]

elif is_bowl and not is_all:
    inn_  = len(bowl_df)
    tbl   = int(bowl_df["Balls_Bowled"].sum()) if "Balls_Bowled" in bowl_df.columns else 0
    ovs   = f"{tbl//6}.{tbl%6}"
    twk   = int(bowl_df["Wickets"].sum())
    rc    = int(bowl_df["Runs_Conceded"].sum()) if "Runs_Conceded" in bowl_df.columns else 0
    ba    = round(rc/max(twk,1), 2)
    eco   = round(bowl_df["Economy"].replace(0,float("nan")).mean(), 2)
    bsr_v = round(tbl/max(twk,1), 1)
    if twk > 0:
        bfr = bowl_df.sort_values(["Wickets","Runs_Conceded"],
                                  ascending=[False,True]).iloc[0]
        bfig= f"{int(bfr['Wickets'])}/{int(bfr.get('Runs_Conceded',0))}"
    else: bfig="0/0"
    kpis = [
        (str(inn_),  "Matches",    scope_lbl,""),
        (ovs,        "Overs",      "Bowled",  ""),
        (str(twk),   "Wickets",    "Total",   "hl"),
        (str(eco),   "Economy",    "RPO",     ""),
        (str(ba),    "Bowl Avg",   "Runs/wkt",""),
        (str(bsr_v), "Bowl SR",    "Balls/wkt",""),
        (bfig,       "Best Figs",  "Wkts/Runs","hl"),
    ]
else:
    bi   = len(bat_df)
    tr2  = int(bat_df["Runs"].sum()) if not bat_df.empty else 0
    no2  = (bat_df.get("Not_out",pd.Series(0,index=bat_df.index))
            if not bat_df.empty else pd.Series([0]))
    av2  = calc_avg(bat_df["Runs"] if not bat_df.empty else pd.Series([0]),
                    no2, max(bi,1))
    sr2  = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(),1) if not bat_df.empty else 0
    fi2  = int(bat_df["50s"].sum()) if not bat_df.empty and "50s" in bat_df.columns else 0
    hu2  = int(bat_df["100s"].sum()) if not bat_df.empty and "100s" in bat_df.columns else 0
    tw2  = int(bowl_df["Wickets"].sum()) if not bowl_df.empty else 0
    ec2  = round(bowl_df["Economy"].replace(0,float("nan")).mean(),2) if not bowl_df.empty and "Economy" in bowl_df.columns else 0
    tb2  = int(bowl_df["Balls_Bowled"].sum()) if not bowl_df.empty and "Balls_Bowled" in bowl_df.columns else 0
    ov2  = f"{tb2//6}.{tb2%6}"
    kpis = [
        (str(bi),       "Innings",   scope_lbl, ""),
        (str(tr2),      "Runs",      "Batting",  "hl"),
        (str(av2),      "Bat Avg",   "Runs/dis", ""),
        (str(sr2),      "SR",        "Avg SR",   ""),
        (f"{fi2}/{hu2}","50s/100s", "Milestones","hl-g"),
        (str(tw2),      "Wickets",  "Bowling",   "hl"),
        (str(ec2),      "Economy",  "RPO",       ""),
        (ov2,           "Overs",    "Bowled",    ""),
    ]

kcols = st.columns(len(kpis))
for i,(v,l,s,var) in enumerate(kpis):
    with kcols[i]:
        st.markdown(kcard(v,l,s,var), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  SECTION 3 — CHARTS
# ══════════════════════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown('<div class="sec-title"><span class="dot"></span>'
            'Strength &amp; Weakness Charts</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
#  BATSMAN  
# ─────────────────────────────────────────────────────────────
if is_bat and not is_all:

    # ══ ROW 1 ════════════════════════════════════════════════
    r1a, r1b, r1c = st.columns(3, gap="large")

    # Chart 1 — Wagon Wheel
    with r1a:
        
        pch(make_wagon_wheel(bat_df, sel_player, h=420))
        

    # Chart 2 — Dismissal Length Weakness (Diverging Bar)
    with r1b:
        
        dl_v = bat_df[~bat_df["Dismissed_Length"].astype(str).isin(
            ["Not Applicable","0","nan",""])]
        if not dl_v.empty:
            dlg = dl_v["Dismissed_Length"].value_counts().reset_index()
            dlg.columns = ["Length", "Dismissals"]
            dlg = dlg.sort_values("Dismissals", ascending=True)
            mx   = dlg["Dismissals"].max()
            mn   = dlg["Dismissals"].min()
            clrs_= [GREEN if d==mn else (RED if d==mx else AMBER)
                    for d in dlg["Dismissals"]]
            fig2 = go.Figure(go.Bar(
                y=dlg["Length"], x=dlg["Dismissals"], orientation="h",
                marker=dict(color=clrs_, line=dict(color="white",width=1.5)),
                text=dlg["Dismissals"], textposition="outside",
                textfont=dict(family="Poppins",size=12),
            ))
            pchart(fig2, 420, "Weakness by Ball Length (Diverging)")
            fig2.update_xaxes(title_text="Number of Dismissals")
            pch(fig2)
        else:
            st.markdown('<div class="no-data">No dismissal length data.</div>',
                        unsafe_allow_html=True)
        

    # Chart 3 — Weakness vs Bowler Type (Spider / Radar)
    with r1c:
        
        da_v = bat_df[~bat_df["Dismissed_against"].astype(str).isin(
            ["Not Applicable","0","nan",""])]
        if not da_v.empty:
            dag = da_v["Dismissed_against"].value_counts().reset_index()
            dag.columns = ["Bowler", "Times"]
            cats_= list(dag["Bowler"]) + [dag["Bowler"].iloc[0]]
            vals_= list(dag["Times"])  + [dag["Times"].iloc[0]]
            fig3 = go.Figure()
            fig3.add_trace(go.Scatterpolar(
                r=vals_, theta=cats_, fill="toself",
                line=dict(color=RED, width=2),
                fillcolor="rgba(192,57,43,0.18)",
                marker=dict(color=RED, size=7),
                name="Dismissals",
            ))
            fig3.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", height=420,
                margin=dict(l=20,r=20,t=44,b=10),
                title=dict(text="Weakness vs Bowler Type (Spider)",
                    font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
                polar=dict(
                    radialaxis=dict(visible=True, gridcolor="#f0f3f8",
                        linecolor=BORDER, tickfont=dict(size=7)),
                    angularaxis=dict(
                        tickfont=dict(family="Inter",size=8,color=NAVY))),
                legend=dict(bgcolor="rgba(255,255,255,.9)",
                    bordercolor=BORDER, borderwidth=1, font=dict(size=10)),
            )
            pch(fig3)
        else:
            st.markdown('<div class="no-data">No dismissal-against data.</div>',
                        unsafe_allow_html=True)
        

    # ══ ROW 2 ════════════════════════════════════════════════
    r2a, r2b, r2c = st.columns(3, gap="large")

    # Chart 4 — Dot Ball Pressure Gauge
    with r2a:
        
        dots_g = int(bat_df["Dots"].sum()) if "Dots" in bat_df.columns else 0
        bf_g   = int(bat_df["Balls_faced"].sum()) if "Balls_faced" in bat_df.columns else 1
        dp_g   = round(dots_g / max(bf_g, 1) * 100, 1)
        fig4   = go.Figure(go.Indicator(
            mode="gauge+number",
            value=dp_g,
            number=dict(suffix="%",
                font=dict(family="Poppins", size=34, color=NAVY)),
            gauge=dict(
                axis=dict(range=[0,100],
                    tickfont=dict(size=9), tickcolor=BORDER),
                bar=dict(color=NAVY, thickness=0.32),
                bgcolor="#f7f9fc",
                borderwidth=1, bordercolor=BORDER,
                steps=[
                    dict(range=[0,30],  color="rgba(30,122,52,.15)"),
                    dict(range=[30,50], color="rgba(232,160,32,.15)"),
                    dict(range=[50,100],color="rgba(192,57,43,.15)"),
                ],
                threshold=dict(
                    line=dict(color=AMBER, width=3),
                    thickness=0.72, value=40),
            ),
        ))
        fig4.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", height=420,
            margin=dict(l=20,r=20,t=44,b=40),
            title=dict(text="Dot Ball Pressure Gauge",
                font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
        )
        fig4.add_annotation(
            x=0.5, y=0.08,
            text=f"{dots_g} dots from {bf_g} balls",
            showarrow=False,
            font=dict(family="Inter",size=10,color="#6B7A8D"),
            xanchor="center")
        fig4.add_annotation(
            x=0.5, y=-0.02,
            text=(" High pressure" if dp_g>45 else "Under control"),
            showarrow=False,
            font=dict(family="Inter",size=10,
                color=RED if dp_g>45 else GREEN),
            xanchor="center")
        pch(fig4)
        

    # Chart 5 — Avg & SR: 1st vs 2nd Innings (Grouped Bar)
    with r2b:
        
        i1b = bat_df[bat_df["Innings_no"]==1]
        i2b = bat_df[bat_df["Innings_no"]==2]
        a1b = calc_avg(i1b["Runs"],
            i1b.get("Not_out",pd.Series(0,index=i1b.index)), len(i1b)) if not i1b.empty else 0
        a2b = calc_avg(i2b["Runs"],
            i2b.get("Not_out",pd.Series(0,index=i2b.index)), len(i2b)) if not i2b.empty else 0
        s1b = round(i1b["Strike_Rate"].replace(0,float("nan")).mean(),1) if not i1b.empty else 0
        s2b = round(i2b["Strike_Rate"].replace(0,float("nan")).mean(),1) if not i2b.empty else 0
        fig5 = go.Figure()
        fig5.add_trace(go.Bar(
            name="Average",
            x=["1st Innings","2nd Innings"], y=[a1b, a2b],
            marker=dict(
                color=[GREEN if a1b>=a2b else STEEL,
                       GREEN if a2b>a1b  else STEEL],
                line=dict(color="white",width=1.5)),
            text=[a1b, a2b], textposition="outside",
            textfont=dict(family="Poppins",size=11),
        ))
        fig5.add_trace(go.Bar(
            name="Strike Rate",
            x=["1st Innings","2nd Innings"], y=[s1b, s2b],
            marker=dict(color=["rgba(232,160,32,0.8)", "rgba(232,160,32,0.8)"],
            line=dict(color="white",width=1.5)),
            text=[s1b, s2b], textposition="outside",
            textfont=dict(family="Poppins",size=11),
        ))
        fig5.update_layout(barmode="group")
        pchart(fig5, 420, "Average & Strike Rate — 1st vs 2nd Innings")
        pch(fig5)
        

    # Chart 6 — Home vs Away Slope Chart
    with r2c:
        
        home_b = bat_df[bat_df["Country"]=="India"] if "Country" in bat_df.columns else pd.DataFrame()
        away_b = bat_df[bat_df["Country"]!="India"] if "Country" in bat_df.columns else pd.DataFrame()
        h_avg  = calc_avg(home_b["Runs"],home_b.get("Not_out",pd.Series(0,index=home_b.index)),len(home_b)) if not home_b.empty else 0
        a_avg  = calc_avg(away_b["Runs"],away_b.get("Not_out",pd.Series(0,index=away_b.index)),len(away_b)) if not away_b.empty else 0
        h_sr   = round(home_b["Strike_Rate"].replace(0,float("nan")).mean(),1) if not home_b.empty else 0
        a_sr   = round(away_b["Strike_Rate"].replace(0,float("nan")).mean(),1) if not away_b.empty else 0
        fig6   = go.Figure()
        for metric, hv, av, hcol, acol in [
            ("Average",     h_avg, a_avg,
             GREEN if h_avg>=a_avg else RED, GREEN if a_avg>h_avg else RED),
            ("Strike Rate", h_sr,  a_sr,
             AMBER if h_sr>=a_sr else STEEL, AMBER if a_sr>h_sr else STEEL),
        ]:
            fig6.add_trace(go.Scatter(
                x=["Home (India)","Away"],
                y=[hv, av],
                mode="lines+markers+text",
                name=metric,
                line=dict(width=2.5,
                    color=GREEN if hv>=av else RED),
                marker=dict(size=13,
                    color=[hcol, acol],
                    line=dict(color="white",width=2)),
                text=[str(hv), str(av)],
                textposition=["top left","top right"],
                textfont=dict(family="Poppins",size=11),
            ))
        pchart(fig6, 420, "Home vs Away Performance — Slope Chart")
        fig6.update_xaxes(tickfont=dict(size=11,family="Poppins"))
        pch(fig6)
        

# ─────────────────────────────────────────────────────────────
#  BOWLER  —  Row 1 × 3 + Row 2 × 2 = 5 charts
# ─────────────────────────────────────────────────────────────
elif is_bowl and not is_all:

    # ══ ROW 1 ════════════════════════════════════════════════
    r1a, r1b, r1c = st.columns(3, gap="large")

    # Chart 1 — Wickets by Length (Radial Bar)
    with r1a:
        
        fd1  = int(bowl_df["Full_Dismissals"].sum()) if "Full_Dismissals" in bowl_df.columns else 0
        gd1  = int(get_col(bowl_df,"Good_Dissmisals","Good_Dismissals").sum())
        sd1  = int(get_col(bowl_df,"Short_Dissmisals","Short_Dismissals").sum())
        wv1  = [fd1, gd1, sd1]
        cr1  = [GREEN if w==max(wv1) else (RED if w==min(wv1) else AMBER)
                for w in wv1]
        fig1 = go.Figure(go.Barpolar(
            r=wv1,
            theta=["Full Length","Good Length","Short Pitch"],
            width=[0.85, 0.85, 0.85],
            marker=dict(color=cr1, line=dict(color="white",width=2.5)),
            hovertemplate="%{theta}: %{r} wickets<extra></extra>",
        ))
        fig1.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", height=380,
            margin=dict(l=10,r=10,t=44,b=10),
            title=dict(text="Wickets by Ball Length (Radial)",
                font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
            polar=dict(
                radialaxis=dict(visible=True, gridcolor="#f0f3f8",
                    linecolor=BORDER, tickfont=dict(size=8)),
                angularaxis=dict(
                    tickfont=dict(family="Poppins",size=11,color=NAVY))),
        )
        pch(fig1)
        

    # Chart 2 — RHB vs LHB Butterfly Chart
    with r1b:
        
        rhb_r = int(bowl_df["Runs_RHB"].sum()) if "Runs_RHB" in bowl_df.columns else 0
        lhb_r = int(bowl_df["Runs_LHB"].sum()) if "Runs_LHB" in bowl_df.columns else 0
        rhb_w = int(bowl_df["Wicket_RHB"].sum()) if "Wicket_RHB" in bowl_df.columns else 0
        lhb_w = int(bowl_df["Wicket_LHB"].sum()) if "Wicket_LHB" in bowl_df.columns else 0
        cats2 = ["Runs Conceded","Wickets Taken"]
        fig2  = go.Figure()
        fig2.add_trace(go.Bar(
            name="vs RHB",
            y=cats2, x=[-rhb_r, -rhb_w],
            orientation="h",
            marker=dict(color=NAVY, line=dict(color="white",width=1.5)),
            text=[rhb_r, rhb_w], textposition="outside",
            textfont=dict(family="Poppins",size=11,color=NAVY),
        ))
        fig2.add_trace(go.Bar(
            name="vs LHB",
            y=cats2, x=[lhb_r, lhb_w],
            orientation="h",
            marker=dict(color=AMBER, line=dict(color="white",width=1.5)),
            text=[lhb_r, lhb_w], textposition="outside",
            textfont=dict(family="Poppins",size=11,color=AMBER),
        ))
        fig2.update_layout(barmode="overlay")
        pchart(fig2, 380, "RHB vs LHB — Butterfly Chart")
        fig2.update_xaxes(
            title_text=" RHB  |  LHB ",
            zeroline=True, zerolinecolor=BORDER, zerolinewidth=2,
            tickvals=[-rhb_r,-rhb_w,0,lhb_r,lhb_w],
            ticktext=[str(rhb_r),str(rhb_w),"0",str(lhb_r),str(lhb_w)],
        )
        pch(fig2)
        

    # Chart 3 — Economy 1st vs 2nd Innings (Grouped Bar)
    with r1c:
        
        eco_all = bowl_df["Economy"].replace(0,float("nan")).dropna()
        half_   = max(len(eco_all)//2, 1)
        eco1_   = round(eco_all.iloc[:half_].mean(), 2)
        eco2_   = round(eco_all.iloc[half_:].mean(), 2)
        wk1_    = int(bowl_df["Wickets"].iloc[:half_].sum())
        wk2_    = int(bowl_df["Wickets"].iloc[half_:].sum())
        fig3    = go.Figure()
        fig3.add_trace(go.Bar(
            name="Economy",
            x=["Early Matches","Recent Matches"],
            y=[eco1_, eco2_],
            marker=dict(
                color=[GREEN if eco1_<=eco2_ else RED,
                       GREEN if eco2_<eco1_  else RED],
                line=dict(color="white",width=1.5)),
            text=[eco1_, eco2_], textposition="outside",
            textfont=dict(family="Poppins",size=11),
        ))
        fig3.add_trace(go.Bar(
            name="Wickets",
            x=["Early Matches","Recent Matches"],
            y=[wk1_, wk2_],
            marker=dict(
            color=["rgba(232,160,32,0.8)", "rgba(232,160,32,0.8)"],
            line=dict(color="white",width=1.5)),
            text=[wk1_, wk2_], textposition="outside",
            textfont=dict(family="Poppins",size=11),
        ))
        fig3.update_layout(barmode="group")
        pchart(fig3, 380, "Economy & Wickets — Match Period Split")
        pch(fig3)
        

    # ══ ROW 2 ════════════════════════════════════════════════
    r2a, r2b = st.columns(2, gap="large")

    # Chart 4 — Bowler Strength Radar
    with r2a:
        
        tb4  = int(bowl_df["Balls_Bowled"].sum()) if "Balls_Bowled" in bowl_df.columns else 0
        tw4  = int(bowl_df["Wickets"].sum())
        db4  = int(get_col(bowl_df,"Dot_Balls").sum())
        ext4 = int(get_col(bowl_df,"Extras").sum())
        eco4 = round(bowl_df["Economy"].replace(0,float("nan")).mean(), 2)
        # normalise 0–100
        dot4  = round(db4/max(tb4,1)*100, 1)
        rctr4 = max(0, round((1 - eco4/10)*100, 1))
        ectr4 = max(0, min(100, round(100 - ext4/max(tb4,1)*200, 1)))
        wkr4  = min(100, round(tw4/max(len(bowl_df),1)*50, 1))
        cats4 = ["Dot Ball %","Run Control","Extras Control","Wicket Rate"]
        vals4 = [dot4, rctr4, ectr4, wkr4]
        # close polygon
        c4_c  = cats4 + [cats4[0]]
        v4_c  = vals4 + [vals4[0]]
        fig4  = go.Figure()
        fig4.add_trace(go.Scatterpolar(
            r=v4_c, theta=c4_c, fill="toself",
            name="Bowling Strength",
            line=dict(color=GREEN, width=2.5),
            fillcolor="rgba(30,122,52,.18)",
            marker=dict(color=GREEN, size=9,
                line=dict(color="white",width=2)),
        ))
        fig4.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", height=380,
            margin=dict(l=20,r=20,t=44,b=10),
            title=dict(text="Bowler Strength Radar",
                font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
            polar=dict(
                radialaxis=dict(range=[0,100], visible=True,
                    gridcolor="#f0f3f8", linecolor=BORDER,
                    tickfont=dict(size=7)),
                angularaxis=dict(
                    tickfont=dict(family="Inter",size=9,color=NAVY))),
            legend=dict(bgcolor="rgba(255,255,255,.9)",
                bordercolor=BORDER,font=dict(size=10)),
        )
        pch(fig4)
        

    # Chart 5 — Home vs Away Bowling Slope
    with r2b:
        
        hbw = bowl_df[bowl_df["Country"]=="India"] if "Country" in bowl_df.columns else pd.DataFrame()
        abw = bowl_df[bowl_df["Country"]!="India"] if "Country" in bowl_df.columns else pd.DataFrame()
        he  = round(hbw["Economy"].replace(0,float("nan")).mean(),2) if not hbw.empty else 0
        ae  = round(abw["Economy"].replace(0,float("nan")).mean(),2) if not abw.empty else 0
        hw  = int(hbw["Wickets"].sum()) if not hbw.empty else 0
        aw  = int(abw["Wickets"].sum()) if not abw.empty else 0
        fig5 = go.Figure()
        for lbl, hv, av, lower_better in [
            ("Economy", he, ae, True),
            ("Wickets", hw, aw, False)
        ]:
            good_h = (hv<=av) if lower_better else (hv>=av)
            fig5.add_trace(go.Scatter(
                x=["Home (India)","Away"],
                y=[hv, av],
                mode="lines+markers+text",
                name=lbl,
                line=dict(width=2.5,
                    color=GREEN if good_h else RED),
                marker=dict(size=13,
                    color=[GREEN if good_h else RED,
                           GREEN if not good_h else RED],
                    line=dict(color="white",width=2)),
                text=[str(hv), str(av)],
                textposition=["top left","top right"],
                textfont=dict(family="Poppins",size=11),
            ))
        pchart(fig5, 380, "Home vs Away Bowling — Slope Chart")
        fig5.update_xaxes(tickfont=dict(size=11,family="Poppins"))
        pch(fig5)
        

# ─────────────────────────────────────────────────────────────
#  ALL-ROUNDER 
# ─────────────────────────────────────────────────────────────
else:
    r1a, r1b = st.columns(2, gap="large")

    # Chart 1 — Skill Balance Radar
    with r1a:
        
        av_r  = calc_avg(bat_df["Runs"] if not bat_df.empty else pd.Series([0]),
                         bat_df.get("Not_out",pd.Series(0,index=bat_df.index)) if not bat_df.empty else pd.Series([0]),
                         max(len(bat_df),1))
        sr_r  = round(bat_df["Strike_Rate"].replace(0,float("nan")).mean(),1) if not bat_df.empty else 0
        tw_r  = int(bowl_df["Wickets"].sum()) if not bowl_df.empty else 0
        ec_r  = round(bowl_df["Economy"].replace(0,float("nan")).mean(),2) if not bowl_df.empty else 0
        tb_r  = int(bowl_df["Balls_Bowled"].sum()) if not bowl_df.empty and "Balls_Bowled" in bowl_df.columns else 0
        db_r  = int(get_col(bowl_df,"Dot_Balls").sum()) if not bowl_df.empty else 0
        dp_r  = round(db_r/max(tb_r,1)*100,1)
        # normalise
        av_n  = min(av_r/100*100, 100)
        sr_n  = min(sr_r/200*100, 100)
        wk_n  = min(tw_r/20*100,  100)
        ec_n  = max(0, (10-ec_r)/10*100)
        cats_r= ["Bat Average","Strike Rate","Wickets","Economy (inv)","Dot Ball %"]
        vals_r= [av_n, sr_n, wk_n, ec_n, dp_r]
        c_rc  = cats_r + [cats_r[0]]
        v_rc  = vals_r + [vals_r[0]]
        figR  = go.Figure()
        figR.add_trace(go.Scatterpolar(
            r=v_rc, theta=c_rc, fill="toself",
            name="Skill Balance",
            line=dict(color=NAVY, width=2.5),
            fillcolor="rgba(27,58,107,.18)",
            marker=dict(color=NAVY,size=9,line=dict(color="white",width=2)),
        ))
        figR.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", height=400,
            margin=dict(l=20,r=20,t=44,b=10),
            title=dict(text="All-Rounder Skill Balance (Radar)",
                font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
            polar=dict(
                radialaxis=dict(range=[0,100],visible=True,
                    gridcolor="#f0f3f8",linecolor=BORDER,
                    tickfont=dict(size=7)),
                angularaxis=dict(
                    tickfont=dict(family="Inter",size=9,color=NAVY))),
            legend=dict(bgcolor="rgba(255,255,255,.9)",
                bordercolor=BORDER,font=dict(size=10)),
        )
        pch(figR)
        

    # Chart 2 — Wagon Wheel
    with r1b:
        
        if not bat_df.empty:
            pch(make_wagon_wheel(bat_df, sel_player, h=400))
        else:
            st.markdown('<div class="no-data">No batting zone data.</div>',
                        unsafe_allow_html=True)
        

    r2a, r2b = st.columns(2, gap="large")

    # Chart 3 — Dismissal Length Weakness
    with r2a:
        
        dl_v = bat_df[~bat_df["Dismissed_Length"].astype(str).isin(
            ["Not Applicable","0","nan",""])]
        if not dl_v.empty:
            dlg = dl_v["Dismissed_Length"].value_counts().reset_index()
            dlg.columns = ["Length", "Dismissals"]
            dlg = dlg.sort_values("Dismissals", ascending=True)
            mx   = dlg["Dismissals"].max()
            mn   = dlg["Dismissals"].min()
            clrs_= [GREEN if d==mn else (RED if d==mx else AMBER)
                    for d in dlg["Dismissals"]]
            fig2 = go.Figure(go.Bar(
                y=dlg["Length"], x=dlg["Dismissals"], orientation="h",
                marker=dict(color=clrs_, line=dict(color="white",width=1.5)),
                text=dlg["Dismissals"], textposition="outside",
                textfont=dict(family="Poppins",size=12),
            ))
            pchart(fig2, 420, "Weakness by Ball Length (Diverging)")
            fig2.update_xaxes(title_text="Number of Dismissals")
            pch(fig2)
        else:
            st.markdown('<div class="no-data">No dismissal length data.</div>',
                        unsafe_allow_html=True)
        
        

    # Chart 4 — Bowling: Radial by Length
    with r2b:
        
        fd4a = int(bowl_df["Full_Dismissals"].sum()) if not bowl_df.empty and "Full_Dismissals" in bowl_df.columns else 0
        gd4a = int(get_col(bowl_df,"Good_Dissmisals","Good_Dismissals").sum()) if not bowl_df.empty else 0
        sd4a = int(get_col(bowl_df,"Short_Dissmisals","Short_Dismissals").sum()) if not bowl_df.empty else 0
        wv4a = [fd4a, gd4a, sd4a]
        cr4a = [GREEN if w==max(wv4a) else (RED if w==min(wv4a) else AMBER)
                for w in wv4a]
        figB = go.Figure(go.Barpolar(
            r=wv4a,
            theta=["Full Length","Good Length","Short Pitch"],
            width=[0.85,0.85,0.85],
            marker=dict(color=cr4a, line=dict(color="white",width=2.5)),
        ))
        figB.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", height=400,
            margin=dict(l=10,r=10,t=44,b=10),
            title=dict(text="Bowling: Wickets by Length (Radial)",
                font=dict(family="Poppins",size=13,color=NAVY),x=0.02),
            polar=dict(
                radialaxis=dict(visible=True,gridcolor="#f0f3f8",
                    linecolor=BORDER,tickfont=dict(size=8)),
                angularaxis=dict(
                    tickfont=dict(family="Poppins",size=11,color=NAVY))),
        )
        pch(figB)
        

# ══════════════════════════════════════════════════════════════
#  ACTION BAR + FOOTER
# ══════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="action-bar">
  <div style="display:flex;align-items:center;gap:12px;">
    <a class="act-btn act-report" href="/generate_report" target="_self">
      &nbsp;Generate Report
    </a>
    <span style="color:#8899aa;font-size:.78rem;">
      {sel_player} &nbsp;·&nbsp; {scope_lbl} &nbsp;·&nbsp; {total_m} matches
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
  Cricket Player Performance Analysis Dashboard &nbsp;|&nbsp;Built with Streamlit
</div>
""", unsafe_allow_html=True)