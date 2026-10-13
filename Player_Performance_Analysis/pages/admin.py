import streamlit as st
import pandas as pd
import openpyxl
import base64, os

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Admin Panel | Cricket Analytics",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
#  PATHS
# ─────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR    = os.path.dirname(BASE_DIR)
IMG_DIR     = os.path.join(ROOT_DIR, "images")
DATASET_DIR = os.path.join(ROOT_DIR, "Dataset")

def img_b64(fname):
    for d in [IMG_DIR, os.path.join(ROOT_DIR, "images"), BASE_DIR]:
        p = os.path.join(d, fname)
        if os.path.exists(p):
            ext  = fname.rsplit(".", 1)[-1].lower()
            mime = "jpeg" if ext in ("jpg", "jpeg") else ext
            with open(p, "rb") as f:
                return f"data:image/{mime};base64,{base64.b64encode(f.read()).decode()}"
    return ""

def _find(fname):
    for d in [DATASET_DIR, ROOT_DIR, BASE_DIR]:
        p = os.path.join(d, fname)
        if os.path.exists(p): return p
    return os.path.join(DATASET_DIR, fname)

MASTER_PATH = _find("player_master.csv")
STATS_PATH  = _find("player_stats2.xlsx")
hero_img    = img_b64("cricket_ds.png")
player_icon = img_b64("profile_icon.png")
stats_icon  = img_b64("stats_icon.png")

ADMIN_CREDS = {
    "admin":      "cricket@2024",
    "superadmin": "admin@123",
}
# ─────────────────────────────────────────────
#  LIVE DATA
# ─────────────────────────────────────────────
@st.cache_data
def load_stats():
    try:
        master = pd.read_csv(MASTER_PATH)
        master.columns = master.columns.str.strip()
    except Exception:
        master = pd.DataFrame(columns=["PlayerName", "Role"])
    player_rows, player_mids, role_map = {}, {}, {}
    unique_mids = set()
    try:
        wb = openpyxl.load_workbook(STATS_PATH)
        for sh in wb.sheetnames:
            ws   = wb[sh]
            rows = list(ws.iter_rows(values_only=True))
            n    = max(len(rows) - 1, 0)
            player_rows[sh] = n
            if n > 0:
                cols = [str(c).strip() if c else f"_c{i}"
                        for i, c in enumerate(rows[0])]
                df = pd.DataFrame(rows[1:], columns=cols)
                mids = sorted(df["Match_id"].dropna().astype(str).unique().tolist()
                              ) if "Match_id" in cols else []
                player_mids[sh] = mids
                unique_mids.update(mids)
            else:
                player_mids[sh] = []
            prow = master[master["PlayerName"].str.strip() == sh]
            role_map[sh] = str(prow.iloc[0]["Role"]) if not prow.empty else "Batsman"
    except Exception:
        pass
    bat_count  = int((master["Role"] == "Batsman").sum())     if not master.empty else 0
    bowl_count = int((master["Role"] == "Bowler").sum())      if not master.empty else 0
    ar_count   = int((master["Role"] == "All-Rounder").sum()) if not master.empty else 0
    return dict(
        master=master, player_rows=player_rows, player_mids=player_mids,
        unique_mids=sorted(unique_mids), unique_match_count=len(unique_mids),
        total_stat_rows=sum(player_rows.values()),
        players_with_data=sum(1 for v in player_rows.values() if v > 0),
        total_registered=len(player_rows), role_map=role_map,
        bat_count=bat_count, bowl_count=bowl_count, ar_count=ar_count,
    )

try:
    SD = load_stats()
except Exception:
    SD = dict(master=pd.DataFrame(), player_rows={}, player_mids={},
              unique_mids=[], unique_match_count=0, total_stat_rows=0,
              players_with_data=0, total_registered=0, role_map={},
              bat_count=0, bowl_count=0, ar_count=0)

# ─────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────
if "admin_logged_in" not in st.session_state:
    st.session_state.admin_logged_in = False
if "admin_user" not in st.session_state:
    st.session_state.admin_user = ""

def do_logout():
    st.session_state.admin_logged_in = False
    st.session_state.admin_user = ""
    st.rerun()

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700;800&display=swap');

/* ── Reset — exact home.py ── */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background: #ffffff !important;
}
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stSidebar"] { display: none !important; }

.block-container,
[data-testid="stAppViewBlockContainer"],
[data-testid="block-container"] {
    padding: 0 60px 60px 60px !important;
    max-width: 1560px !important;
    margin: 0 auto !important;
    background: #fff !important;
}
[data-testid="stAppViewBlockContainer"] > div:first-child {
    padding-top: 0 !important;
}

/* ── NAVBAR — exact home.py white style ── */
.top-header {
    background: #ffffff;
    border-bottom: 2px solid #1B3A6B;
    padding: 0 60px;
    height: 64px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    position: fixed;
    width: 100%;
    top: 0;
    z-index: 999;
    box-shadow: 0 2px 12px rgba(27,58,107,0.08);
    margin: 0 -60px 0 -60px;
}
.header-brand {
    font-family: 'Poppins', sans-serif;
    font-size: 1.25rem;
    font-weight: 700;
    color: #1B3A6B;
    letter-spacing: 0.3px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.header-brand span { color: #E8A020; }
.header-nav { display: flex; gap: 8px; align-items: center; }

/* home.py .h-btn — solid navy */
.h-btn {
    background: #1B3A6B;
    color: #ffffff !important;
    border: none;
    padding: 7px 20px;
    border-radius: 6px;
    font-family: 'Inter', sans-serif;
    font-size: 0.82rem;
    font-weight: 600;
    cursor: pointer;
    letter-spacing: 0.3px;
    text-decoration: none !important;
    transition: background 0.2s;
    display: inline-block;
}
.h-btn:hover { background: #243f7a; color: #fff !important; text-decoration: none !important; }

/* home.py .h-btn-ghost — outline navy */
.h-btn-ghost {
    background: transparent;
    color: #1B3A6B !important;
    border: 1.5px solid #1B3A6B;
    padding: 6px 18px;
    border-radius: 6px;
    font-family: 'Inter', sans-serif;
    font-size: 0.82rem;
    font-weight: 600;
    cursor: pointer;
    letter-spacing: 0.3px;
    text-decoration: none !important;
    transition: background 0.2s;
    display: inline-block;
}
.h-btn-ghost:hover { background: #f0f4fb; text-decoration: none !important; }

/* logout — red, same pill shape */
.h-btn-logout {
    background: #C0392B;
    color: #ffffff !important;
    border: none;
    padding: 7px 20px;
    border-radius: 6px;
    font-family: 'Inter', sans-serif;
    font-size: 0.82rem;
    font-weight: 600;
    cursor: pointer;
    letter-spacing: 0.3px;
    text-decoration: none !important;
    transition: background 0.2s;
    display: inline-flex;
    align-items: center;
    gap: 6px;
}
.h-btn-logout:hover { background: #a02220; color: #fff !important; text-decoration: none !important; }

/* admin username pill */
.admin-pill {
    background: #EBF3FF;
    color: #1B3A6B;
    padding: 3px 12px;
    border-radius: 20px;
    font-size: 0.7rem;
    font-weight: 700;
    border: 1px solid #c0d0eb;
    letter-spacing: 0.4px;
}

/* ── Streamlit widget overrides ── */
div[data-testid="stButton"] > button {
    background: #1B3A6B !important;
    color: #ffffff !important;
    font-family: 'Poppins', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.9rem !important;
    letter-spacing: 0.8px !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 13px 0 !important;
    width: 100% !important;
    transition: background 0.2s !important;
}
div[data-testid="stButton"] > button:hover { background: #243f7a !important; }
.logout-wrap div[data-testid="stButton"] > button {
    background: #C0392B !important;
    padding: 8px 0 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.3px !important;
}
.logout-wrap div[data-testid="stButton"] > button:hover { background: #a02220 !important; }

div[data-testid="stTextInput"] input {
    border: 1.5px solid #c8d0e0 !important;
    border-radius: 8px !important;
    font-size: 0.92rem !important;
    color: #1B3A6B !important;
    background: #fff !important;
    padding: 11px 14px !important;
}
div[data-testid="stTextInput"] input:focus {
    border-color: #1B3A6B !important;
    box-shadow: 0 0 0 3px rgba(27,58,107,0.1) !important;
}
div[data-testid="stAlert"] { border-radius: 8px !important; font-size: 0.84rem !important; }

/* ══════════════════════
   LOGIN PAGE
══════════════════════ */
.login-page-bg {
    min-height: calc(100vh - 64px);
    background: linear-gradient(145deg, #f0f4fb 0%, #EBF3FF 40%, #f7f9fc 100%);
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 48px 24px;
    margin: 0 -60px;
}
.login-card {
    background: #ffffff;
    border: 1.5px solid #dde3ee;
    border-radius: 16px;
    padding: 52px 56px 44px 56px;
    width: 100%;
    max-width: 460px;
    box-shadow: 0 8px 32px rgba(27,58,107,0.10);
    position: relative;
    overflow: hidden;
}
.login-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: #1B3A6B;
}
.login-logo {
    width: 64px; height: 64px;
    background: #EBF3FF;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 2rem;
    margin: 0 auto 20px auto;
    border: 1.5px solid #c0d0eb;
}
.login-title {
    font-family: 'Poppins', sans-serif;
    font-size: 1.55rem;
    font-weight: 800;
    color: #1B3A6B;
    text-align: center;
    margin-top: 200px;
}
.login-subtitle {
    font-size: 0.84rem;
    color: #6B7A8D;
    text-align: center;
    line-height: 1.6;
    margin-bottom: 32px;
}
.login-field-lbl {
    display: block;
    font-size: 0.75rem;
    font-weight: 700;
    color: #1B3A6B;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    margin-bottom: 6px;
}
.login-secure {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    font-size: 0.72rem;
    color: #8899aa;
    margin-top: 16px;
    letter-spacing: 0.3px;
}

/* ══════════════════════════════════════
   DASHBOARD — all from home.py
══════════════════════════════════════ */

/* HERO — exact home.py */
.hero-wrap {
    padding: 52px 0 40px 0;
    display: flex;
    align-items: center;
    gap: 48px;
}
.hero-left { flex: 1.2; }
.hero-right { flex: 1; }
.hero-eyebrow {
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: #E8A020;
    margin-bottom: 12px;
}
.hero-title {
    font-family: 'Poppins', sans-serif;
    font-size: 2.4rem;
    font-weight: 800;
    color: #1B3A6B;
    line-height: 1.2;
    margin-bottom: 18px;
}
.hero-title span { color: #E8A020; }
.hero-desc {
    font-size: 0.95rem;
    color: #5a6a80;
    line-height: 1.8;
    margin-bottom: 28px;
    max-width: 520px;
}
.hero-stats { display: flex; gap: 32px; margin-top: 8px; }
.hero-stat-val {
    font-family: 'Poppins', sans-serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: #1B3A6B;
    line-height: 1;
}
.hero-stat-lbl {
    font-size: 0.72rem;
    color: #8899aa;
    font-weight: 500;
    margin-top: 3px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* SECTION DIVIDER — exact home.py */
.sec-divider {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, #1B3A6B 0%, #e8edf5 60%);
    margin: 8px 0 36px 0;
}
.sec-header { display: flex; align-items: center; gap: 10px; margin-bottom: 24px; }
.sec-header-title {
    font-family: 'Poppins', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: #1B3A6B;
}
.sec-header-line { flex: 1; height: 1px; background: #e8edf5; }

/* STATS STRIP — exact home.py */
.stats-strip {
    background: #1B3A6B;
    border-radius: 12px;
    padding: 24px 36px;
    display: flex;
    justify-content: space-around;
    align-items: center;
    margin: 12px 0 40px 0;
    gap: 20px;
}
.strip-item { text-align: center; }
.strip-val {
    font-family: 'Poppins', sans-serif;
    font-size: 1.6rem;
    font-weight: 700;
    color: #ffffff;
}
.strip-lbl {
    font-size: 0.72rem;
    color: rgba(255,255,255,0.65);
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-top: 3px;
    font-weight: 500;
}
.strip-divider { width: 1px; height: 40px; background: rgba(255,255,255,0.2); }

/* ANALYSIS CARDS — exact home.py (cards-grid used for admin action cards) */
.cards-grid-2 {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 18px;
    margin-bottom: 20px;
}
.analysis-card {
    background: #ffffff;
    border: 1.5px solid #dde3ee;
    border-radius: 14px;
    padding: 28px 16px 22px 16px;
    text-align: center;
    text-decoration: none !important;
    transition: all 0.25s ease;
    display: block;
    cursor: pointer;
    position: relative;
    overflow: hidden;
}
.analysis-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    background: #1B3A6B;
    transform: scaleX(0);
    transition: transform 0.25s ease;
}
.analysis-card:hover::before { transform: scaleX(1); }
.analysis-card:hover {
    border-color: #1B3A6B;
    box-shadow: 0 8px 28px rgba(27,58,107,0.14);
    transform: translateY(-4px);
    text-decoration: none !important;
}
.analysis-card.amber-card::before { background: #E8A020; }
.analysis-card.amber-card:hover { border-color: #E8A020; }

.card-icon-wrap {
    width: 60px;
    height: 60px;
    background: #EBF3FF;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin: 0 auto 16px auto;
    transition: background 0.25s;
}
.analysis-card:hover .card-icon-wrap { background: #1B3A6B; }
.analysis-card:hover .card-icon-wrap img { filter: brightness(0) invert(1); }
.card-icon-wrap img {
    width: 30px;
    height: 30px;
    object-fit: contain;
    transition: filter 0.25s;
}

.card-title {
    font-family: 'Poppins', sans-serif;
    font-size: 1.0rem;
    font-weight: 700;
    color: #1B3A6B;
    margin-bottom: 8px;
    text-decoration: none !important;
}
.analysis-card.amber-card .card-title { color: #B06010; }
.card-desc { font-size: 0.8rem; color: #8899aa; line-height: 1.6; margin-bottom: 18px; }
.card-badge {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    background: #EBF3FF;
    border: 1px solid #c0d0eb;
    border-radius: 20px;
    padding: 3px 13px;
    font-size: 0.72rem;
    font-weight: 700;
    color: #1B3A6B;
    margin-bottom: 18px;
}
.analysis-card.amber-card .card-badge {
    background: #FFF8EE;
    border-color: #f0d090;
    color: #B06010;
}
.card-cta {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 9px 24px;
    border-radius: 7px;
    background: #1B3A6B;
    color: #fff !important;
    font-family: 'Poppins', sans-serif;
    font-weight: 700;
    font-size: 0.82rem;
    text-decoration: none !important;
    letter-spacing: 0.3px;
}
.analysis-card.amber-card .card-cta { background: #E8A020; }

/* KPI OVERVIEW — 4 cols */
.kpi-overview-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 18px;
    margin-bottom: 12px;
}
.kpi-overview-card {
    background: #ffffff;
    border: 1.5px solid #dde3ee;
    border-radius: 14px;
    padding: 24px 16px 20px 16px;
    text-align: center;
    box-shadow: 0 2px 10px rgba(27,58,107,0.06);
    position: relative;
    overflow: hidden;
    transition: all 0.25s ease;
}
.kpi-overview-card:hover {
    border-color: #1B3A6B;
    box-shadow: 0 8px 24px rgba(27,58,107,0.12);
    transform: translateY(-3px);
}
.kpi-overview-card::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
    background: #1B3A6B;
}
.kpi-overview-card.amber::after { background: #E8A020; }
.kpi-overview-card.green::after  { background: #1E7A34; }
.kpi-icon-wrap {
    width: 52px; height: 52px;
    background: #EBF3FF;
    border-radius: 13px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin: 0 auto 12px auto;
    font-size: 1.4rem;
}
.kpi-overview-card.amber .kpi-icon-wrap { background: #FFF8EE; }
.kpi-overview-card.green  .kpi-icon-wrap { background: #EEF8F0; }
.kpi-overview-val {
    font-family: 'Poppins', sans-serif;
    font-size: 1.75rem;
    font-weight: 700;
    color: #1B3A6B;
    line-height: 1;
    margin-bottom: 5px;
}
.kpi-overview-lbl {
    font-size: 0.72rem;
    color: #8899aa;
    font-weight: 600;
    letter-spacing: 0.6px;
    text-transform: uppercase;
}
.kpi-overview-sub { font-size: 0.7rem; color: #6B7A8D; margin-top: 4px; }

/* DATA KPI — 5 cols */
.data-kpi-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 16px;
    margin-bottom: 12px;
}
.data-kpi-card {
    background: #ffffff;
    border: 1.5px solid #dde3ee;
    border-radius: 12px;
    padding: 20px 12px 16px 12px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(27,58,107,0.05);
    position: relative;
    overflow: hidden;
    transition: box-shadow 0.2s, transform 0.2s;
}
.data-kpi-card:hover { box-shadow: 0 6px 20px rgba(27,58,107,0.12); transform: translateY(-2px); }
.data-kpi-card::after { content: ''; position: absolute; bottom: 0; left: 0; right: 0; height: 3px; background: #1B3A6B; }
.data-kpi-card.dkc-amber::after { background: #E8A020; }
.data-kpi-card.dkc-green::after { background: #1E7A34; }
.data-kpi-card.dkc-steel::after { background: #4A6FA5; }
.data-kpi-card.dkc-red::after   { background: #C0392B; }
.dkc-icon  { font-size: 1.3rem; margin-bottom: 8px; }
.dkc-val   { font-family: 'Poppins', sans-serif; font-size: 1.6rem; font-weight: 700; color: #1B3A6B; line-height: 1; margin-bottom: 5px; }
.dkc-lbl   { font-size: 0.68rem; color: #8899aa; font-weight: 600; letter-spacing: 0.5px; text-transform: uppercase; }
.dkc-sub   { font-size: 0.68rem; color: #6B7A8D; margin-top: 3px; }

/* MATCH BREAKDOWN NOTE */
.breakdown-note {
    background: #f7f9fc;
    border: 1px solid #dde3ee;
    border-radius: 8px;
    padding: 12px 18px;
    margin-bottom: 14px;
    font-size: 0.82rem;
    color: #4A6FA5;
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
}

/* HOW IT WORKS — exact home.py .how-grid / .how-card */
.how-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 20px;
    margin-bottom: 48px;
}
.how-card {
    background: #f7f9fc;
    border: 1px solid #dde3ee;
    border-radius: 12px;
    padding: 24px 20px;
}
.how-step {
    width: 32px; height: 32px;
    background: #1B3A6B;
    color: white;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-family: 'Poppins', sans-serif;
    font-weight: 700;
    font-size: 0.9rem;
    margin-bottom: 14px;
}
.how-title { font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 0.92rem; color: #1B3A6B; margin-bottom: 8px; }
.how-desc  { font-size: 0.8rem; color: #6B7A8D; line-height: 1.6; }

/* FOOTER — exact home.py */
.footer {
    background:#1B3A6B; color:rgba(255,255,255,.7);
    text-align:center; padding:18px 60px;
    font-size:.8rem; margin:0 -60px -60px -60px;
    letter-spacing:.3px;
}
.footer span { color:#E8A020; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  NAVBAR
# ─────────────────────────────────────────────
if st.session_state.admin_logged_in:
        st.markdown(f"""
        <div class="top-header">
          <div class="header-brand">
            Cricket Player<span> Performance Analysis</span>
          </div>
          <div class="header-nav">
            <a class="h-btn" href="/userhome" target="_self">Home</a>
            <a class="h-btn" href="/"         target="_self">Back</a>
            <a class="h-btn" href="/userhome" target="_self">LOG OUT</a>
          </div>
        </div>
        """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="top-header">
      <div class="header-brand">
        Cricket Player<span> Performance Analysis</span>
      </div>
      <div class="header-nav">
        <a class="h-btn" href="/userhome" target="_self">Home</a>
        <a class="h-btn" href="/"         target="_self">Back</a>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════
#  ①  LOGIN PAGE
# ══════════════════════════════════════════
if not st.session_state.admin_logged_in:
    _, lc, _ = st.columns([1, 1.5, 1])
    with lc:
        st.markdown("""
          <div class="login-title">Admin Login</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)

        st.markdown(
            '<span class="login-field-lbl">&nbsp;Username</span>',
            unsafe_allow_html=True)
        username = st.text_input("Username", placeholder="Enter your username",
                                  label_visibility="collapsed", key="lu")

        st.markdown(
            '<span class="login-field-lbl" '
            'style="display:block;margin-top:14px;">&nbsp;Password</span>',
            unsafe_allow_html=True)
        password = st.text_input("Password", placeholder="Enter your password",
                                  type="password",
                                  label_visibility="collapsed", key="lp")

        st.markdown('<div style="height:10px;"></div>', unsafe_allow_html=True)
        login_btn = st.button("Sign In", use_container_width=True)

        if login_btn:
            u = username.strip()
            if not u:
                st.error("Please enter your username.")
            elif not password:
                st.error("Please enter your password.")
            elif u in ADMIN_CREDS and ADMIN_CREDS[u] == password:
                st.session_state.admin_logged_in = True
                st.session_state.admin_user = u
                st.rerun()
            else:
                st.error("Invalid credentials. Please try again.")

        st.markdown("""
        <div class="login-secure">
          &nbsp;Secured access &nbsp;·&nbsp; Authorised personnel only
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)   # end .login-page-bg

    st.stop()


# ══════════════════════════════════════════
#  ②  ADMIN DASHBOARD
# ══════════════════════════════════════════

# ── HERO (mirrors home.py structure exactly) ──
col_hero_l, col_hero_r = st.columns([1.3, 1], gap="large")

with col_hero_l:
    st.markdown(f"""
    <div style="padding: 48px 0 32px 0;">
      <div class="hero-eyebrow">Admin Control Panel</div>
      <div class="hero-title">
        Manage Cricket<br>Data <span>Like a Pro</span>
      </div>
      <p class="hero-desc">
        This admin panel allows administrators to manage cricket player data,
        match statistics, and update the analytics dataset used across the platform.
        Add new players, upload match-by-match statistics, and maintain the full
        dataset — powering every insight for analysts and team strategists.
      </p>
      <div class="hero-stats">
        <div>
          <div class="hero-stat-val">{SD['total_registered']}</div>
          <div class="hero-stat-lbl">Players in DB</div>
        </div>
        <div style="width:1px;background:#dde3ee;"></div>
        <div>
          <div class="hero-stat-val">{SD['unique_match_count']}</div>
          <div class="hero-stat-lbl">Matches Recorded</div>
        </div>
        <div style="width:1px;background:#dde3ee;"></div>
        <div>
          <div class="hero-stat-val">{SD['players_with_data']}</div>
          <div class="hero-stat-lbl">Players Active</div>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

with col_hero_r:
    st.markdown("<div style='padding-top:48px;'>", unsafe_allow_html=True)
    if hero_img:
        st.markdown(f"""
        <div style="border-radius:16px;overflow:hidden;
                    border:1.5px solid #dde3ee;
                    box-shadow:0 8px 32px rgba(27,58,107,0.10);">
          <img src="{hero_img}" style="width:100%;display:block;">
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="background:#f7f9fc;border:1.5px solid #dde3ee;
                    border-radius:16px;height:280px;
                    display:flex;align-items:center;
                    justify-content:center;font-size:4rem;"></div>
        """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ── STATS STRIP ───────────────────────────
st.markdown(f"""
<div class="stats-strip">
  <div class="strip-item">
    <div class="strip-val">ODI</div><div class="strip-lbl">Format</div>
  </div>
  <div class="strip-divider"></div>
  <div class="strip-item">
    <div class="strip-val">India</div><div class="strip-lbl">Team</div>
  </div>
  <div class="strip-divider"></div>
  <div class="strip-item">
    <div class="strip-val">After 2019 CWC</div><div class="strip-lbl">Data Period</div>
  </div>
  <div class="strip-divider"></div>
  <div class="strip-item">
    <div class="strip-val">{SD['total_registered']}</div>
    <div class="strip-lbl">Players in DB</div>
  </div>
  <div class="strip-divider"></div>
  <div class="strip-item">
    <div class="strip-val">{SD['unique_match_count']}</div>
    <div class="strip-lbl">Unique Matches</div>
  </div>
</div>
""", unsafe_allow_html=True)
# ── ADMIN ACTION CARDS (2-col, exact analysis-card style) ──
st.markdown("""
<div class="sec-header">
  <div class="sec-header-title">Admin Actions</div>
  <div class="sec-header-line"></div>
</div>
""", unsafe_allow_html=True)


def icon_html(b64_src, emoji_fallback):
    if b64_src:
        return f'<img src="{b64_src}" alt="">'
    return f'<span style="font-size:1.9rem;line-height:0.8;">{emoji_fallback}</span>'

st.markdown(f"""
<div class="cards-grid-2">

  <a href="/add_player" target="_self" class="analysis-card">
    <div class="card-icon-wrap">{icon_html(player_icon, "")}</div>
    <div class="card-title">Add Player</div>
    <div class="card-desc">
      Register a new player into the dataset — full name, role
      (Batsman / Bowler / All-Rounder), batting &amp; bowling style,
      country, date of birth, and debut year.
    </div>
    <div class="card-badge"> {SD['total_registered']} Players Registered</div>
  </a>

  <a href="/add_stats" target="_self" class="analysis-card amber-card">
    <div class="card-icon-wrap">{icon_html(stats_icon, "")}</div>
    <div class="card-title">Add Match Stats</div>
    <div class="card-desc">
      Enter match-by-match player statistics — batting figures
      (runs, balls, zones), bowling figures (wickets, economy, dot balls)
      and match result for any registered player.
    </div>
    <div class="card-badge"> {SD['unique_match_count']} Matches · {SD['players_with_data']} Players Active</div>
  </a>

</div>
""", unsafe_allow_html=True)

st.markdown('<div class="sec-divider"></div>', unsafe_allow_html=True)

# ── HOW IT WORKS ──
st.markdown("""
<div class="sec-header">
  <div class="sec-header-title">How the Admin Panel Works</div>
  <div class="sec-header-line"></div>
</div>
<div class="how-grid">
  <div class="how-card">
    <div class="how-step">1</div>
    <div class="how-title">Admin Logs In</div>
    <div class="how-desc">
      Admin securely signs in using credentials.
      The session is protected and clears on logout,
      keeping dataset management fully secure.
    </div>
  </div>
  <div class="how-card">
    <div class="how-step">2</div>
    <div class="how-title">Add New Players</div>
    <div class="how-desc">
      Register players with full profiles — name, role,
      batting/bowling style, country, date of birth and
      debut year. Creates their master entry.
    </div>
  </div>
  <div class="how-card">
    <div class="how-step">3</div>
    <div class="how-title">Enter Match Stats</div>
    <div class="how-desc">
      Select a player and enter match-level figures —
      date, opponent, venue, result and role-specific
      stats. Validations ensure accuracy.
    </div>
  </div>
  <div class="how-card">
    <div class="how-step">4</div>
    <div class="how-title">Analytics Update</div>
    <div class="how-desc">
      Every saved record instantly powers all live dashboards —
      overall performance, recent form, opponent analysis,
      strength &amp; weakness, and comparison reports.
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── FOOTER — exact home.py ─────────────────
st.markdown("""
<div class="footer">
  Cricket Player Performance Analysis &nbsp;|&nbsp; Built with Streamlit
</div>
""", unsafe_allow_html=True)