import streamlit as st
import pandas as pd
from datetime import date
import os
import re

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Admin | Cricket Analysis",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# PATHS
# ─────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATASET_DIR = os.path.join(BASE_DIR, "Dataset")

MASTER_PATH = os.path.join(DATASET_DIR, "player_master.csv")
STATS_PATH  = os.path.join(DATASET_DIR, "player_stats2.xlsx")

def _find(fname):
    for d in [DATASET_DIR, BASE_DIR]:
        p = os.path.join(d, fname)
        if os.path.exists(p):
            return p
    return os.path.join(DATASET_DIR, fname)

MASTER_PATH = _find("player_master.csv")
STATS_PATH  = _find("player_stats2.xlsx")

# ─────────────────────────────────────────────
#  GLOBAL CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700;800&display=swap');

html,body,[class*="css"]{font-family:'Inter',sans-serif;background:#fff !important;}
#MainMenu,footer,header{visibility:hidden;}
[data-testid="stSidebar"]{display:none;}

.block-container,
[data-testid="stAppViewBlockContainer"],
[data-testid="block-container"]{
    padding:0 60px 60px 60px !important;
    max-width:1560px !important;
    margin:0 auto !important;
    background:#fff !important;
}
[data-testid="stAppViewBlockContainer"]>div:first-child{padding-top:0 !important;}

/* ── NAVBAR ── */
.top-header{
    background:#fff;border-bottom:2px solid #1B3A6B;
    padding:0 60px;height:64px;
    display:flex;align-items:center;justify-content:space-between;
    position:fixed; width:100%; top:0;z-index:999;
    box-shadow:0 2px 12px rgba(27,58,107,.08);
    margin:0 -60px 0 -60px;
}
.header-brand{font-family:'Poppins',sans-serif;font-size:1.2rem;font-weight:700;color:#1B3A6B;}
.header-brand span{color:#E8A020;}
.header-nav{display:flex;gap:8px;align-items:center;}
.h-btn{background:#1B3A6B;color:#fff !important;border:none;padding:7px 18px;border-radius:6px;
       font-size:.82rem;font-weight:600;text-decoration:none !important;display:inline-block;transition:background .2s;}
.h-btn:hover{background:#243f7a;}
.h-btn-ghost{background:transparent;color:#1B3A6B !important;border:1.5px solid #1B3A6B;
             padding:6px 16px;border-radius:6px;font-size:.82rem;font-weight:600;
             text-decoration:none !important;display:inline-block;transition:background .2s;}
.h-btn-ghost:hover{background:#f0f4fb;}

/* ── PAGE TITLE ── */
.page-title-bar{background:#f7f9fc;border-left:4px solid #1B3A6B;
    padding:14px 60px;margin:40px -60px 28px -60px;}
.page-title-bar h2{font-family:'Poppins',sans-serif;font-size:1.15rem;font-weight:700;color:#1B3A6B;margin:0;}
.page-title-bar p{color:#6B7A8D;font-size:.8rem;margin:3px 0 0 0;}

/* ── FILTER / PLAYER SELECTOR PANEL ── */
.filter-wrap{background:#f7f9fc;border:1px solid #dde3ee;border-radius:10px;
             padding:22px 28px;margin-bottom:26px;}
.filter-label{font-size:.75rem;font-weight:600;color:#1B3A6B;
              letter-spacing:.8px;text-transform:uppercase;margin-bottom:6px;}

/* ── STREAMLIT WIDGET OVERRIDES ── */
.stSelectbox>div>div{background:#fff !important;border:1.5px solid #c8d0e0 !important;
    border-radius:7px !important;color:#1B3A6B !important;font-size:.88rem !important;}
.stSelectbox>div>div:focus-within{border-color:#1B3A6B !important;
    box-shadow:0 0 0 3px rgba(27,58,107,.1) !important;}

div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input{
    border:1.5px solid #c8d0e0 !important;border-radius:7px !important;
    font-size:.88rem !important;color:#1B3A6B !important;background:#fff !important;}
div[data-testid="stTextInput"] input:focus,
div[data-testid="stNumberInput"] input:focus{
    border-color:#1B3A6B !important;box-shadow:0 0 0 3px rgba(27,58,107,.08) !important;}
div[data-testid="stTextInput"] input:disabled,
div[data-testid="stNumberInput"] input:disabled{
    background:#f7f9fc !important;color:#8899aa !important;border-color:#e8edf5 !important;}

div[data-testid="stDateInput"] input{
    border:1.5px solid #c8d0e0 !important;border-radius:7px !important;
    font-size:.88rem !important;color:#1B3A6B !important;}

div[data-testid="stButton"]>button{background:#1B3A6B !important;color:#fff !important;
    font-family:'Poppins',sans-serif !important;font-weight:700 !important;
    font-size:.9rem !important;letter-spacing:1px !important;border:none !important;
    border-radius:7px !important;padding:12px 0 !important;width:100% !important;}
div[data-testid="stButton"]>button:hover{background:#243f7a !important;}
div[data-testid="stButton"]>button:disabled{
    background:#dde3ee !important;color:#8899aa !important;cursor:not-allowed !important;}

/* ── SECTION TITLE ── */
.sec-title{font-family:'Poppins',sans-serif;font-size:1rem;font-weight:700;color:#1B3A6B;
           margin:0 0 14px 0;padding-bottom:8px;border-bottom:2px solid #e8edf5;
           display:flex;align-items:center;gap:8px;}
.sec-title .dot{width:8px;height:8px;background:#1B3A6B;border-radius:50%;display:inline-block;}
.sec-title .dot-amber{width:8px;height:8px;background:#E8A020;border-radius:50%;display:inline-block;}

/* ── SUB-SECTION LABEL ── */
.subsec-label{font-family:'Poppins',sans-serif;font-size:.78rem;font-weight:700;
    color:#4A6FA5;text-transform:uppercase;letter-spacing:.6px;
    padding:6px 12px 6px 12px;border-left:3px solid #E8A020;
    background:#fffaf0;border-radius:0 6px 6px 0;margin:20px 0 12px 0;
    display:flex;align-items:center;gap:8px;}

/* ── FORM CARD ── */
.form-card{background:#fff;border:1px solid #dde3ee;border-radius:12px;
    padding:24px 24px 20px 24px;box-shadow:0 2px 10px rgba(27,58,107,.05);
    margin-bottom:20px;position:relative;overflow:hidden;}
.form-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
    background:linear-gradient(90deg,#1B3A6B 0%,#E8A020 100%);}

/* ── PLAYER INFO PILL ── */
.player-pill{display:flex;align-items:center;gap:12px;
    background:#EBF3FF;border:1px solid #c0d0eb;border-radius:8px;
    padding:10px 16px;margin-top:4px;}
.player-pill-name{font-family:'Poppins',sans-serif;font-size:.92rem;font-weight:700;color:#1B3A6B;}
.player-role-badge{display:inline-block;padding:2px 12px;border-radius:20px;font-size:.68rem;
    font-weight:700;letter-spacing:.5px;text-transform:uppercase;}
.badge-batsman{background:#EBF3FF;color:#1B3A6B;border:1px solid #c0d0eb;}
.badge-bowler{background:#FFF0E8;color:#C0501A;border:1px solid #f0c0a0;}
.badge-allrounder{background:#F0F8EE;color:#2A7A2A;border:1px solid #a8d4a0;}

/* ── LAST MATCHES TABLE ── */
.recent-table-wrap{
    background:#fff;border:1px solid #dde3ee;border-radius:12px;
    overflow:hidden;box-shadow:0 2px 10px rgba(27,58,107,.05);
    margin-bottom:6px;position:relative;}
.recent-table-wrap::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;
    background:linear-gradient(90deg,#1B3A6B 0%,#E8A020 100%);}
.recent-table-inner{overflow-x:auto;padding:0;}
.recent-table-wrap table{width:100%;border-collapse:collapse;font-size:.8rem;}
.recent-table-wrap thead tr{background:#1B3A6B;}
.recent-table-wrap thead th{
    color:#fff;font-family:'Poppins',sans-serif;font-size:.7rem;font-weight:700;
    letter-spacing:.6px;text-transform:uppercase;padding:10px 14px;
    text-align:center;white-space:nowrap;border:none;}
.recent-table-wrap thead th:first-child{text-align:left;}
.recent-table-wrap tbody tr{border-bottom:1px solid #f0f3f8;}
.recent-table-wrap tbody tr:last-child{border-bottom:none;}
.recent-table-wrap tbody tr:hover{background:#f7f9fc;}
.recent-table-wrap tbody td{
    padding:9px 14px;color:#2c3e50;font-size:.8rem;
    text-align:center;white-space:nowrap;}
.recent-table-wrap tbody td:first-child{text-align:left;font-weight:600;color:#1B3A6B;}
.result-won{display:inline-block;padding:2px 10px;background:#EAF6EE;color:#1E7A34;
    border:1px solid #b8dfc4;border-radius:12px;font-size:.68rem;font-weight:700;}
.result-loss{display:inline-block;padding:2px 10px;background:#FDE8E8;color:#C0392B;
    border:1px solid #f0b8b8;border-radius:12px;font-size:.68rem;font-weight:700;}
.result-other{display:inline-block;padding:2px 10px;background:#f7f9fc;color:#6B7A8D;
    border:1px solid #dde3ee;border-radius:12px;font-size:.68rem;font-weight:700;}
.no-history{text-align:center;padding:32px 20px;color:#8899aa;font-size:.85rem;
    background:#f7f9fc;border-radius:0 0 12px 12px;}
.recent-table-header{
    padding:14px 20px 12px 20px;
    border-bottom:1px solid #e8edf5;
    display:flex;align-items:center;justify-content:space-between;}
.recent-table-title{font-family:'Poppins',sans-serif;font-size:.88rem;font-weight:700;
    color:#1B3A6B;}
.recent-table-count{font-size:.72rem;color:#8899aa;font-weight:500;}

/* ── MATCH TOTALS BOX ── */
.match-totals-box{
    background:linear-gradient(135deg,#1B3A6B 0%,#243f7a 100%);
    border-radius:10px;padding:18px 20px 16px 20px;margin:18px 0 6px 0;}
.mtb-title{font-family:'Poppins',sans-serif;font-size:.72rem;font-weight:700;
    color:rgba(255,255,255,.6);letter-spacing:.9px;text-transform:uppercase;
    margin-bottom:12px;}
.mtb-grid{display:flex;flex-wrap:wrap;gap:8px;}
.mtb-item{background:rgba(255,255,255,.10);border:1px solid rgba(255,255,255,.15);
    border-radius:8px;padding:10px 14px;text-align:center;flex:1;min-width:90px;}
.mtb-item.hl .mtb-val{color:#E8A020;}
.mtb-val{font-family:'Poppins',sans-serif;font-size:1.1rem;font-weight:700;
    color:#fff;line-height:1;margin-bottom:4px;}
.mtb-lbl{font-size:.62rem;color:rgba(255,255,255,.55);font-weight:600;
    letter-spacing:.5px;text-transform:uppercase;}

/* ── KPI CARDS ── */
.kpi-card{background:#fff;border:1px solid #dde3ee;border-radius:10px;
    padding:16px 12px 12px 12px;text-align:center;
    box-shadow:0 2px 8px rgba(27,58,107,.05);
    position:relative;overflow:hidden;}
.kpi-card::after{content:'';position:absolute;bottom:0;left:0;right:0;
    height:3px;background:#1B3A6B;}
.kpi-card.hl::after{background:#E8A020;}
.kpi-card.hl-g::after{background:#1E7A34;}
.kpi-val{font-family:'Poppins',sans-serif;font-size:1.55rem;font-weight:700;
    color:#1B3A6B;line-height:1;margin-bottom:6px;}
.kpi-lbl{font-size:.68rem;color:#8899aa;font-weight:600;letter-spacing:.6px;text-transform:uppercase;}

/* ── INSIGHT / ERROR / SUCCESS CARDS ── */
.insight-card{background:#f7f9fc;border:1px solid #dde3ee;border-left:4px solid #E8A020;
    border-radius:8px;padding:12px 16px;display:flex;align-items:flex-start;
    gap:12px;margin-bottom:9px;}
.insight-icon{font-size:1.1rem;flex-shrink:0;margin-top:2px;}
.insight-text{font-size:.84rem;color:#2c3e50;line-height:1.5;}
.insight-text strong{color:#1B3A6B;}

.error-card{background:#FDE8E8;border:1px solid #f0b8b8;border-left:4px solid #C0392B;
    border-radius:8px;padding:12px 16px;margin-bottom:8px;
    font-size:.83rem;color:#8B1A1A;display:flex;align-items:center;gap:10px;}
.success-card{background:#EAF6EE;border:1px solid #b8dfc4;border-left:4px solid #1E7A34;
    border-radius:8px;padding:13px 16px;margin:12px 0;
    font-size:.85rem;color:#1E5228;font-weight:600;display:flex;align-items:center;gap:10px;}

/* ── DIVIDER ── */
.divider{border:none;height:1px;background:#e8edf5;margin:28px 0;}

/* ── NO DATA STATE ── */
.no-data{text-align:center;padding:60px 20px;color:#8899aa;font-size:.9rem;}

/* ── ACTION / SAVE BAR ── */
.action-bar{background:#f7f9fc;border-top:1px solid #dde3ee;
    padding:22px 60px;margin:32px -60px 0 -60px;
    display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap;}
.save-hint{font-size:.8rem;color:#8899aa;line-height:1.6;}
.save-hint strong{color:#1B3A6B;}
.save-status{font-size:.8rem;font-weight:600;}
.save-status.ok{color:#1E7A34;}
.save-status.err{color:#C0392B;}

/* ── FOOTER ── */
.footer{background:#1B3A6B;color:rgba(255,255,255,.7);text-align:center;
    padding:18px 60px;font-size:.8rem;margin:0 -60px -60px -60px;letter-spacing:.3px;}
.footer span{color:#E8A020;font-weight:600;}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data
def load_players():
    df = pd.read_csv(MASTER_PATH)
    df.columns = df.columns.str.strip()
    return df

players_df = load_players()
player_list = ["Select Player"] + players_df["PlayerName"].unique().tolist()


# ─────────────────────────────────────────────
#  LAST 10 MATCHES HELPER
# ─────────────────────────────────────────────
def load_last_10_matches(player, role):
    """
    Load the last 10 match entries for the given player from player_stats2.xlsx.
    Returns a DataFrame (may be empty if no data found).
    """
    if not os.path.exists(STATS_PATH):
        return pd.DataFrame()
    try:
        df = pd.read_excel(STATS_PATH, sheet_name=player)
    except Exception:
        return pd.DataFrame()

    if df.empty:
        return pd.DataFrame()

    # Sort by Match_date descending if the column exists
    if "Match_date" in df.columns:
        df["Match_date"] = pd.to_datetime(df["Match_date"], errors="coerce")
        df = df.sort_values("Match_date", ascending=False)

    return df.head(10)


def render_last_10_table(player, role):
    """
    Render the last 10 match entries for the player as a styled HTML table.
    Columns shown depend on the player's role.
    """
    df = load_last_10_matches(player, role)

    st.markdown(
        '<div class="sec-title"><span class="dot"></span>Last 10 Match Entries</div>',
        unsafe_allow_html=True,
    )

    if df.empty:
        st.markdown(
            '<div class="recent-table-wrap">'
            '<div class="no-history">No match entries found for this player.</div>'
            '</div>',
            unsafe_allow_html=True,
        )
        return

    total = len(df)

    # ── Choose columns to display based on role ──────────────
    base_cols = ["Match_id", "Match_date", "Opponent", "Venue", "Country", "Team Score","Match_result"]

    batting_cols = [
        "Batted_Flag", "Innings_no", "Runs", "Balls_faced",
        "4s", "6s", "Strike_Rate", "50s", "100s", "Not_out",
        "Dismissed_method",
    ]
    bowling_cols = [
        "Bowl_Flag", "Balls_Bowled", "Runs_Conceded",
        "Wickets", "Maidens", "Economy", "Dot_Balls", "Extras",
    ]

    if role == "Batsman":
        desired_cols = base_cols + batting_cols
    elif role == "Bowler":
        desired_cols = base_cols + bowling_cols
    else:  # All-Rounder
        desired_cols = base_cols + batting_cols + bowling_cols

    # Keep only columns that actually exist in the data
    show_cols = [c for c in desired_cols if c in df.columns]
    df_display = df[show_cols].copy()

    # Format date
    if "Match_date" in df_display.columns:
        df_display["Match_date"] = df_display["Match_date"].dt.strftime("%d %b %Y")

    # Format numeric columns to avoid .0 suffixes
    int_cols = [
        "Innings_no", "Runs", "Balls_faced", "4s", "6s", "50s", "100s",
        "Not_out", "Batted_Flag", "Bowl_Flag", "Balls_Bowled",
        "Runs_Conceded", "Wickets", "Maidens", "Dot_Balls", "Extras",
    ]
    for col in int_cols:
        if col in df_display.columns:
            df_display[col] = df_display[col].fillna(0).astype(int)

    float_cols = ["Strike_Rate", "Economy"]
    for col in float_cols:
        if col in df_display.columns:
            df_display[col] = df_display[col].fillna(0.0).round(2)

    # ── Build HTML table ──────────────────────────────────────
    header_html = "".join(
        f'<th>{col.replace("_", " ")}</th>' for col in df_display.columns
    )

    rows_html = ""
    for _, row in df_display.iterrows():
        cells = ""
        for col in df_display.columns:
            val = row[col]
            if col == "Match_result":
                if str(val).strip().lower() == "won":
                    badge = f'<span class="result-won">Won</span>'
                elif str(val).strip().lower() == "loss":
                    badge = f'<span class="result-loss">Loss</span>'
                else:
                    badge = f'<span class="result-other">{val}</span>'
                cells += f"<td>{badge}</td>"
            elif col == "Match_id":
                cells += f"<td>{val}</td>"
            else:
                cells += f"<td>{val}</td>"
        rows_html += f"<tr>{cells}</tr>"

    st.markdown(
        f"""
        <div class="recent-table-wrap">
          <div class="recent-table-header">
            <span class="recent-table-title">Recent Match History</span>
            <span class="recent-table-count">Showing last {total} record(s)</span>
          </div>
          <div class="recent-table-inner">
            <table>
              <thead><tr>{header_html}</tr></thead>
              <tbody>{rows_html}</tbody>
            </table>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
#  NAVBAR
# ─────────────────────────────────────────────
st.markdown("""
<div class="top-header">
  <div class="header-brand">Cricket Player<span> Performance Analysis</span></div>
  <div class="header-nav">
    <a class="h-btn" href="/home" target="_self">User</a>
    <a class="h-btn" href="/" target="_self">Back</a>
  </div>
</div>
""", unsafe_allow_html=True)

# PAGE TITLE BAR
st.markdown("""
<div class="page-title-bar">
  <h2>Admin — Player Stats Entry</h2>
  <p>Enter and save match performance data for players across all formats</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  PLAYER SELECTOR
# ─────────────────────────────────────────────
st.markdown('<div class="filter-label">Select Player</div>', unsafe_allow_html=True)
player_name = st.selectbox("Player", player_list, label_visibility="collapsed")


if player_name != "Select Player":
    prow     = players_df[players_df["PlayerName"] == player_name].iloc[0]
    role_cls = (
        "badge-batsman"    if prow["Role"] == "Batsman"    else
        "badge-bowler"     if prow["Role"] == "Bowler"     else
        "badge-allrounder"
    )
    st.markdown(f"""
    <div class="player-pill">
      <div>
        <div class="player-pill-name">{player_name}</div>
        <span class="player-role-badge {role_cls}">{prow['Role']}</span>
      </div>
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  STOP IF NO PLAYER
# ─────────────────────────────────────────────
if player_name == "Select Player":
    st.markdown("""
    <div class="no-data">
      Select a player above to begin entering match statistics.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

player_row  = players_df[players_df["PlayerName"] == player_name].iloc[0]
player_role = player_row["Role"]


# ─────────────────────────────────────────────
#  LAST 10 MATCHES TABLE
# ─────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)
render_last_10_table(player_name, player_role)


# ═══════════════════════════════════════════════
#  SECTION — MATCH INFORMATION
# ═══════════════════════════════════════════════
st.markdown('<hr class="divider">', unsafe_allow_html=True)
st.markdown(
    '<div class="sec-title"><span class="dot"></span>Match Information</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)
with col1:
    match_id    = st.text_input("Match ID", help="Format: MATCH_000  (e.g. MATCH_001)")
    match_date = st.date_input("Match Date",value=date.today(),max_value=date.today())
    st.markdown('<div class="filter-label"> Select Opponent</div>', unsafe_allow_html=True)

    odi_nations = [
        "Afghanistan","Australia","Bangladesh","England","India","Ireland",
        "New Zealand","Pakistan","South Africa","Sri Lanka","West Indies",
        "Zimbabwe","Netherlands","Scotland","Nepal","Oman",
        "United Arab Emirates","United States","Namibia","Papua New Guinea"
    ]

    opponent = st.selectbox("Opponent", ["All"] + sorted(odi_nations), label_visibility="collapsed")

    match_result = st.selectbox("Match Result", ["Won", "Loss", "No Result", "Tie"])

with col2:
    venue        = st.text_input("Venue")
    country      = st.selectbox("Country", ["All"] + sorted(odi_nations), label_visibility="collapsed")
    total    = st.number_input("Team Score", min_value=0)
    

# Match-level Totals Box
date_disp = match_date.strftime("%d %b %Y") if match_date else "—"

st.markdown(f"""
<div class="match-totals-box">
  <div class="mtb-title">Match Details Summary</div>
  <div class="mtb-grid">
    <div class="mtb-item">
      <div class="mtb-val">{match_id.strip() or "—"}</div>
      <div class="mtb-lbl">Match ID</div>
    </div>
    <div class="mtb-item">
      <div class="mtb-val" style="font-size:.88rem;">{date_disp}</div>
      <div class="mtb-lbl">Date</div>
    </div>
    <div class="mtb-item">
      <div class="mtb-val" style="font-size:.88rem;">{opponent.strip() or "—"}</div>
      <div class="mtb-lbl">Opponent</div>
    </div>
    <div class="mtb-item">
      <div class="mtb-val" style="font-size:.82rem;">{venue.strip() or "—"}</div>
      <div class="mtb-lbl">Venue</div>
    </div>
    <div class="mtb-item">
      <div class="mtb-val" style="font-size:.88rem;">{country.strip() or "—"}</div>
      <div class="mtb-lbl">Country</div>
    </div>
    <div class="mtb-item hl">
      <div class="mtb-val" style="font-size:.88rem;">{match_result}</div>
      <div class="mtb-lbl">Result</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<hr class="divider">', unsafe_allow_html=True)


# ═══════════════════════════════════════════════
#  BATTING FORM
# ═══════════════════════════════════════════════
if "batting_flag" not in st.session_state:
    st.session_state.batting_flag = 0

def batting_form():
    errors   = []
    is_valid = True

    st.markdown(
        '<div class="sec-title"><span class="dot"></span>Batting Performance</div>',
        unsafe_allow_html=True
    )

    batting_flag = st.selectbox("Batting Flag", [0, 1], key="batting_flag",
                                help="1 = Batted this match")
    disabled = batting_flag == 0

    # Scoring Details
    st.markdown('<div class="subsec-label">Scoring Details</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)

    with c1:
        innings_no  = st.selectbox("Innings No", [1, 2], disabled=disabled, key="innings_no")
        balls_faced = st.number_input("Balls Faced", min_value=0, disabled=disabled, key="balls_faced")
        runs        = st.number_input("Runs", min_value=0,
                                      disabled=disabled or balls_faced == 0, key="runs")
        disable_runs  = runs == 0
        enable_50  = runs >= 50 and runs < 100
        enable_100 = runs < 100
        fifty = st.toggle("50s",value=st.session_state.get("fifty", False),disabled=not enable_50,key="fifty")

    with c2:
        dots   = st.number_input("Dot Balls", min_value=0, disabled=disabled, key="dots")
        
        max_fours = runs // 4
        max_sixes = runs // 6
        fours_val = st.session_state.get("fours", 0)
        sixes_val = st.session_state.get("sixes", 0)
        remaining_after_fours = runs - (fours_val * 4)
        remaining_after_sixes = runs - (sixes_val * 6)
        allowed_sixes = remaining_after_fours // 6
        allowed_fours = remaining_after_sixes // 4
        allowed_sixes = max(0, allowed_sixes)
        allowed_fours = max(0, allowed_fours)

        fours = st.number_input("4s",min_value=0,max_value=max_fours,value=fours_val,disabled=runs < 4 or allowed_fours == 0,key="fours")

        sixes = st.number_input("6s",min_value=0,max_value=max_sixes,value=sixes_val,disabled=runs < 6 or allowed_sixes == 0,key="sixes")
        hundred = st.toggle("100s", disabled=enable_100, key="hundred")

    with c3:
        singles_calc = runs - (fours * 4 + sixes * 6)
        singles      = st.number_input("1s + 2s + 3s (Auto)",
                                       value=max(singles_calc, 0), disabled=True)
        strike_rate  = (runs / balls_faced * 100) if balls_faced > 0 else 0.0
        st.number_input("Strike Rate (Auto)", value=round(strike_rate, 2), disabled=True)
        not_out = st.toggle("Not Out", disabled=disabled, key="not_out")
        if balls_faced > 0 and runs > balls_faced * 6:
            st.warning("Impossible runs for given balls")
            is_valid = False
            errors.append(f"Runs cannot exceed {balls_faced * 6} (Balls x 6)")
        score   = singles + (fours * 4) + (sixes * 6)
        if score != runs and not disabled:
            st.warning("Scoring mismatch — 1s + 4s + 6s should equal total runs")
            is_valid = False
            errors.append("Enter valid boundaries and singles")

    # Wagon Wheel & Dismissal
    c4, c5, c6 = st.columns(3)

    with c4:
        zone_vals = {
        "longon": st.session_state.get("longon", 0),
        "longoff": st.session_state.get("longoff", 0),
        "covers": st.session_state.get("covers", 0),
        "point": st.session_state.get("point", 0),
        "thirdman": st.session_state.get("thirdman", 0),
        "fineleg": st.session_state.get("fineleg", 0),
        "square": st.session_state.get("square", 0),
        "midwicket": st.session_state.get("midwicket", 0),
        }
        zone_total = sum(zone_vals.values())
        remaining_runs = max(0, runs - zone_total)
        lock_zones = zone_total >= runs

        longon = st.number_input("Runs Long On",min_value=0,max_value=runs,value=zone_vals["longon"],disabled=disable_runs or (lock_zones and zone_vals["longon"] == 0),key="longon")
        longoff = st.number_input("Runs Long Off",min_value=0,max_value=runs,value=zone_vals["longoff"],disabled=disable_runs or (lock_zones and zone_vals["longoff"] == 0),key="longoff")
        covers = st.number_input("Runs Covers",min_value=0,max_value=runs,value=zone_vals["covers"],disabled=disable_runs or (lock_zones and zone_vals["covers"] == 0),key="covers")
        point = st.number_input("Runs Point",min_value=0,max_value=runs,value=zone_vals["point"],disabled=disable_runs or (lock_zones and zone_vals["point"] == 0),key="point")

    with c5:
        thirdman = st.number_input("Runs Third Man", min_value=0,max_value=runs,value=zone_vals["thirdman"],disabled=disable_runs or (lock_zones and zone_vals["thirdman"] == 0),key="thirdman")
        fineleg = st.number_input("Runs Fine Leg",min_value=0,max_value=runs,value=zone_vals["fineleg"],disabled=disable_runs or (lock_zones and zone_vals["fineleg"] == 0),key="fineleg")
        square = st.number_input("Runs Square",min_value=0,max_value=runs,value=zone_vals["square"],disabled=disable_runs or (lock_zones and zone_vals["square"] == 0),key="square")
        midwicket = st.number_input("Runs Mid Wicket",min_value=0,max_value=runs,value=zone_vals["midwicket"],disabled=disable_runs or (lock_zones and zone_vals["midwicket"] == 0),key="midwicket")

    with c6:
        dismissed_method  = st.selectbox(
            "Dismissal Method",
            ["Not Applicable","Caught","Bowled","LBW","Run Out","Stumped",
             "Hit Wicket","Handling the Ball","Hit the Ball Twice",
             "Obstructing the Field","Retried Out","Timed Out"],
            disabled=disabled or not_out, key="dismissed_method"
        )
        dismissed_length  = st.selectbox(
            "Dismissed Length",
            ["Not applicable","Full","Good","Short"],
            disabled=disabled or not_out, key="dismissed_length"
        )
        dismissed_against = st.text_input("Dismissed Against", disabled=disabled or not_out,key="dismissed_against")

        dismissed_by      = st.selectbox(
            "Dismissed By",
            ["Not Applicable","RA Fast","LA Fast","RA Medium","LA Medium",
             "RA Off-Spin","LA Off-Spin","RA Leg-Break","LA Leg-Break"],
            disabled=disabled or not_out, key="dismissed_by"
        )

        zone_total = (longon + longoff + covers + point +
                      thirdman + fineleg + square + midwicket)
        if not disable_runs and zone_total != runs:
            is_valid = False
            errors.append(f"Zone total ({zone_total}) does not match Runs ({runs})")
        if batting_flag == 1 and not not_out:
            if dismissed_against.strip() == "":
                is_valid = False
                errors.append("Dismissed Against is required when player is out")
        if batting_flag == 1 and not_out:
           if batting_flag == 1 and not_out:
                dismissed_against = "Not Applicable"

    # Batting Summary Box
    if batting_flag == 1:
        bnd_runs = fours * 4 + sixes * 6
        bnd_pct  = round(bnd_runs / max(runs, 1) * 100, 1)
        st.markdown(f"""
        <div class="match-totals-box" style="margin-top:20px;">
          <div class="mtb-title">Batting Summary</div>
          <div class="mtb-grid">
            <div class="mtb-item hl">
              <div class="mtb-val">{int(runs)}</div>
              <div class="mtb-lbl">Runs</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{int(balls_faced)}</div>
              <div class="mtb-lbl">Balls</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{round(strike_rate, 1)}</div>
              <div class="mtb-lbl">Strike Rate</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{int(fours)}/{int(sixes)}</div>
              <div class="mtb-lbl">4s / 6s</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{int(dots)}</div>
              <div class="mtb-lbl">Dots</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{bnd_pct}%</div>
              <div class="mtb-lbl">Boundary %</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{"Yes" if fifty else "No"}</div>
              <div class="mtb-lbl">50s</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{"Yes" if hundred else "No"}</div>
              <div class="mtb-lbl">100s</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{"Yes" if not_out else "No"}</div>
              <div class="mtb-lbl">Not Out</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)


    data = {
        "Batted_Flag":       batting_flag,
        "Innings_no":        innings_no,
        "Runs":              runs,
        "Balls_faced":       balls_faced,
        "Dots":              dots,
        "1s_2s_3s":          singles,
        "4s":                fours,
        "6s":                sixes,
        "50s":               int(fifty),
        "100s":              int(hundred),
        "Not_out":           int(not_out),
        "Strike_Rate":       round(strike_rate, 2),
        "Runs_Longon":       longon,
        "Runs_Longoff":      longoff,
        "Runs_Covers":       covers,
        "Runs_Point":        point,
        "Runs_Thirdman":     thirdman,
        "Runs_Fineleg":      fineleg,
        "Runs_Square":       square,
        "Runs_Midwct":       midwicket,
        "Dismissed_method":  dismissed_method,
        "Dismissed_Length":  dismissed_length,
        "Dismissed_against": "Not Applicable" if not_out else dismissed_against,
        "Dismissed_by":      dismissed_by,
    }
    return data, is_valid, errors


# ═══════════════════════════════════════════════
#  BOWLING FORM
# ═══════════════════════════════════════════════
def bowling_form():
    errors   = []
    is_valid = True

    st.markdown(
        '<div class="sec-title"><span class="dot"></span>Bowling Performance</div>',
        unsafe_allow_html=True
    )

    bowl_flag = st.selectbox("Bowling Flag", [0, 1], key="bowl_flag",
                             help="1 = Bowled this match")
    disabled  = bowl_flag == 0

    # Bowling Figures
    c1, c2, c3 = st.columns(3)

    with c1:
        balls_bowled  = st.number_input("Balls Bowled",  min_value=0, step=1,
                                        disabled=disabled, key="balls_bowled")
        runs_conceded = st.number_input("Runs Conceded", min_value=0,
                                        disabled=disabled, key="runs_conceded")
        dot_balls = st.number_input("Dots Bowled",min_value=0,disabled=disabled or balls_bowled == 0,key="dot_balls")

        auto_maidens = dot_balls // 6

        maidens = st.number_input("Maidens",min_value=0,value=auto_maidens,disabled=dot_balls <= 6,key="maidens")
        extras        = st.number_input("Extras",        min_value=0,
                                        disabled=disabled, key="extras")
        wickets       = st.number_input("Wickets",       min_value=0,
                                        disabled=disabled or (balls_bowled == 0 and extras == 0), key="wickets")
    

    with c2:
        economy = (runs_conceded / (balls_bowled / 6)) if balls_bowled > 0 else 0.0
        st.number_input("Economy (Auto)", value=round(economy, 2), disabled=True)
        balls_rhb = st.number_input("Balls to RHB",min_value=0,max_value=60,disabled=disabled or balls_bowled == 0,key="balls_rhb")
        balls_lhb = st.number_input("Balls to LHB",min_value=0,max_value=60,disabled=disabled or balls_bowled == 0,key="balls_lhb")
        runs_rhb = st.number_input("Runs to RHB",min_value=0,max_value=100,disabled=disabled or balls_bowled == 0,key="runs_rhb")
        runs_lhb = st.number_input("Runs to LHB",min_value=0,max_value=100,disabled=disabled or balls_bowled == 0,key="runs_lhb")
        
    with c3:
        wickt_rhb  = st.number_input("Wickets vs RHB",         min_value=0, max_value=10,
                                     disabled= wickets == 0  , key="wickt_rhb")
        wickt_lhb  = st.number_input("Wickets vs LHB",         min_value=0, max_value=10,
                                     disabled= wickets == 0 != wickets, key="wickt_lhb")
        full_dsml  = st.number_input("Full Length Dismissals",  min_value=0, max_value=10,
                                     disabled=wickets == 0, key="full_dsml")
        good_dsml  = st.number_input("Good Length Dismissals",  min_value=0, max_value=10,
                                     disabled=wickets == 0, key="good_dsml")
        short_dsml = st.number_input("Short Length Dismissals", min_value=0, max_value=10,
                                     disabled= wickets == 0  , key="short_dsml")
        

        if not disabled:
            if balls_bowled != (balls_rhb + balls_lhb):
                is_valid = False
                errors.append("Balls Bowled mismatch(RHB & LHB)")
            if runs_conceded != (runs_rhb + runs_lhb + extras):
                is_valid = False
                errors.append("Runs Conceded mismatch(RHB, LHB & Extras)")
            if wickets != (wickt_rhb + wickt_lhb) == (full_dsml + good_dsml + short_dsml):
                is_valid = False
                errors.append("Wickets mismatch(LHB & LHB)")
            if balls_bowled == 0 and wickets != extras:
                is_valid = False
                errors.append("Wickets cannot exceed Balls Bowled")

    # Bowling Summary Box
    if bowl_flag == 1:
        overs_str = f"{balls_bowled // 6}.{balls_bowled % 6}"
        dot_pct   = round(dot_balls / max(balls_bowled, 1) * 100, 1)
        bowl_sr   = round(balls_bowled / max(wickets, 1), 1) if wickets > 0 else "—"
        st.markdown(f"""
        <div class="match-totals-box" style="margin-top:20px;">
          <div class="mtb-title">Bowling Summary</div>
          <div class="mtb-grid">
            <div class="mtb-item">
              <div class="mtb-val">{overs_str}</div>
              <div class="mtb-lbl">Overs</div>
            </div>
            <div class="mtb-item hl">
              <div class="mtb-val">{int(wickets)}</div>
              <div class="mtb-lbl">Wickets</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{int(runs_conceded)}</div>
              <div class="mtb-lbl">Runs</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{round(economy, 2)}</div>
              <div class="mtb-lbl">Economy</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{int(maidens)}</div>
              <div class="mtb-lbl">Maidens</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{dot_pct}%</div>
              <div class="mtb-lbl">Dot %</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{bowl_sr}</div>
              <div class="mtb-lbl">Bowl SR</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{int(wickt_rhb)}/{int(wickt_lhb)}</div>
              <div class="mtb-lbl">RHB/LHB Wkts</div>
            </div>
            <div class="mtb-item">
              <div class="mtb-val">{int(extras)}</div>
              <div class="mtb-lbl">Extras</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    data = {
        "Bowl_Flag":        bowl_flag,
        "Balls_Bowled":     balls_bowled,
        "Runs_Conceded":    runs_conceded,
        "Maidens":          maidens,
        "Dot_Balls":        dot_balls,
        "Extras":           extras,
        "Economy":          round(economy, 2),
        "Wickets":          wickets,
        "Balls_to_RHB":     balls_rhb,
        "Balls_to_LHB":     balls_lhb,
        "Full_Dismissals":  full_dsml,
        "Good_Dismissals":  good_dsml,
        "Short_Dismissals": short_dsml,
        "Runs_RHB":         runs_rhb,
        "Runs_LHB":         runs_lhb,
        "Wicket_RHB":       wickt_rhb,
        "Wicket_LHB":       wickt_lhb,
    }
    return data, is_valid, errors


# ═══════════════════════════════════════════════
#  ROLE-BASED FORM RENDERING
# ═══════════════════════════════════════════════
can_save   = True
all_errors = []
stats_data = {}

if player_role == "Batsman":
    bat_data, bat_valid, bat_errors = batting_form()
    stats_data.update(bat_data)
    can_save  &= bat_valid
    all_errors.extend(bat_errors)

elif player_role == "Bowler":
    bowl_data, bowl_valid, bowl_errors = bowling_form()
    stats_data.update(bowl_data)
    can_save  &= bowl_valid
    all_errors.extend(bowl_errors)

elif player_role == "All-Rounder":
    bat_data, bat_valid, bat_errors = batting_form()
    stats_data.update(bat_data)
    can_save  &= bat_valid
    all_errors.extend(bat_errors)

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    bowl_data, bowl_valid, bowl_errors = bowling_form()
    stats_data.update(bowl_data)
    can_save  &= bowl_valid
    all_errors.extend(bowl_errors)


# ═══════════════════════════════════════════════
#  CLEAR FORM HELPER
# ═══════════════════════════════════════════════
def clear_form(role):
    batting_keys = [
        "batting_flag","innings_no","runs","balls_faced","fifty","dots",
        "fours","sixes","hundred","not_out","longon","longoff","covers",
        "point","thirdman","fineleg","square","midwicket",
        "dismissed_method","dismissed_length","dismissed_against","dismissed_by",
    ]
    bowling_keys = [
        "bowl_flag","balls_bowled","runs_conceded","maidens","dot_balls","extras",
        "wickets","balls_rhb","balls_lhb","full_dsml","good_dsml","short_dsml",
        "runs_rhb","runs_lhb","wickt_rhb","wickt_lhb",
    ]
    keys_to_clear = (
        batting_keys if role == "Batsman" else
        bowling_keys if role == "Bowler"  else
        batting_keys + bowling_keys
    )
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()


# ═══════════════════════════════════════════════
#  VALIDATIONS
# ═══════════════════════════════════════════════
def valid_match_id(mid):
    return bool(re.match(r"^MATCH_[0-9]{3}$", mid))

def cross_verify_match(file_path, match_id, match_date, opponent, venue, result):
    if not os.path.exists(file_path):
        return True, None

    try:
        xls = pd.ExcelFile(file_path)

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            if "Match_id" not in df.columns:
                continue

            row = df[df["Match_id"].astype(str) == str(match_id)]

            if not row.empty:
                row = row.iloc[0]

                existing_date   = pd.to_datetime(row["Match_date"]).date()
                existing_opp    = str(row["Opponent"]).strip()
                existing_venue  = str(row["Venue"]).strip()
                existing_result = str(row["Match_result"]).strip()

                if (
                    existing_date == match_date
                    and existing_opp.lower() == opponent.lower()
                    and existing_venue.lower() == venue.lower()
                    and existing_result.lower() == result.lower()
                ):
                    return True, sheet
                else:
                    return False, sheet

    except Exception:
        return True, None

    return True, None

def is_duplicate(fp, sheet, mid):
    if not os.path.exists(fp):
        return False
    try:
        df = pd.read_excel(fp, sheet_name=sheet)
        return mid in df.get("Match_id", pd.Series(dtype=str)).astype(str).values
    except Exception:
        return False
    
balls_bowled = stats_data.get("Balls_Bowled", 0)
bowl_flag    = stats_data.get("Bowl_Flag", 0)

if bowl_flag == 1 and balls_bowled == 0:
    can_save = False
    all_errors.append("Enter the bowling stats.")

wickets = stats_data.get("Wickets", 0)
extras = stats_data.get("Extras", 0)

if balls_bowled == 0 and wickets > 0 and extras == 0:
    can_save = False
    all_errors.append(
        "Wicket without a legal delivery requires extras (wide). Please enter Extras."
    )
maidens= stats_data.get("Maidens", 0)
dot_balls = stats_data.get("Dot_Balls", 0)
if maidens != (dot_balls // 6):
    is_valid = False
    all_errors.append("Maidens must equal Dot Balls ÷ 6")

if not valid_match_id(match_id):
    can_save = False
    all_errors.append("Match ID must be in format MATCH_000  (e.g. MATCH_001)")

if any(v.strip() == "" for v in [match_id, opponent, venue, country]):
    can_save = False
    all_errors.append("Please fill in all Match Information fields")

if is_duplicate(STATS_PATH, player_name, match_id):
    can_save = False
    all_errors.append(f"Match ID '{match_id}' already exists for {player_name}")

ok, sheet_found = cross_verify_match(
    STATS_PATH,
    match_id,
    match_date,
    opponent,
    venue,
    match_result
)

if not ok:
    can_save = False
    all_errors.append(
        "This player is not part of this match. Please verify the match details."
    )


# ═══════════════════════════════════════════════
#  ERROR CARDS
# ═══════════════════════════════════════════════
if all_errors:
    st.markdown('<hr class="divider">', unsafe_allow_html=True)
    for err in all_errors:
        st.markdown(
            f'<div class="error-card">&nbsp;{err}</div>',
            unsafe_allow_html=True
        )


# ═══════════════════════════════════════════════
#  ACTION BAR  —  Save button
# ═══════════════════════════════════════════════
status_txt = "All validations passed — ready to save" if can_save else f"{len(all_errors)} error(s) to resolve before saving"
status_cls = "save-status ok" if can_save else "save-status err"

st.markdown('<div class="action-bar">', unsafe_allow_html=True)
bar_l, bar_m, bar_r = st.columns([3, 2.5, 1.5])

with bar_l:
    st.markdown(f"""
    <div class="save-hint">
      <strong>{player_name}</strong> &nbsp;·&nbsp; {player_role}<br>
      Match: <strong>{match_id.strip() or "—"}</strong> &nbsp;·&nbsp;
      vs {opponent.strip() or "—"} &nbsp;·&nbsp; {match_result}
    </div>
    """, unsafe_allow_html=True)

with bar_m:
    st.markdown(f'<div class="{status_cls}" style="padding:8px 0;">{status_txt}</div>',
                unsafe_allow_html=True)

with bar_r:
    save = st.button("Save Stats", disabled=not can_save, use_container_width=True)

st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════
#  SAVE LOGIC
# ═══════════════════════════════════════════════
if save:
    final_data = {
        "Match_id":     match_id,
        "Match_date":   match_date,
        "Opponent":     opponent,
        "Venue":        venue,
        "Country":      country,
        "Match_result": match_result,
        "Team_score" : total,
        **stats_data,
    }
    df_new = pd.DataFrame([final_data])

    file_path = STATS_PATH
    if os.path.exists(file_path):
        try:
            df_existing = pd.read_excel(file_path, sheet_name=player_name)
            if str(match_id) in df_existing["Match_id"].astype(str).values:
                st.error(f"Match ID '{match_id}' already exists for {player_name}.")
                st.stop()
            df_final = pd.concat([df_existing, df_new], ignore_index=True)
        except ValueError:
            df_final = df_new
    else:
        df_final = df_new

    with pd.ExcelWriter(
        file_path,
        engine="openpyxl",
        mode="a" if os.path.exists(file_path) else "w",
        if_sheet_exists="replace",
    ) as writer:
        df_final.to_excel(writer, sheet_name=player_name, index=False)
    

    st.markdown("""
    <div class="success-card">&nbsp; Player stats saved successfully!</div>
    """, unsafe_allow_html=True)
    clear_form(player_role)


# ═══════════════════════════════════════════════
#  FOOTER
# ═══════════════════════════════════════════════
st.markdown("""
<div class="footer">
  Cricket Player Performance Analysis Dashboard &nbsp;|&nbsp;
  Built with Streamlit
</div>
""", unsafe_allow_html=True)