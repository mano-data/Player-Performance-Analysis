import streamlit as st
import pandas as pd
import os
from datetime import date

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Add Player | Cricket Analysis",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
#  PATHS
# ─────────────────────────────────────────────
BASE_DIR  = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DATA_DIR  = os.path.join(BASE_DIR, "Dataset")
IMAGE_DIR = os.path.join(BASE_DIR, "images")

file_path = os.path.join(DATA_DIR, "player_master.csv")

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
    position:fixed; width:100%;top:0;z-index:999;
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

/* ── SUB-SECTION LABEL ── */
.subsec-label{font-family:'Poppins',sans-serif;font-size:.78rem;font-weight:700;
    color:#4A6FA5;text-transform:uppercase;letter-spacing:.6px;
    padding:6px 12px 6px 12px;border-left:3px solid #E8A020;
    background:#fffaf0;border-radius:0 6px 6px 0;margin:20px 0 12px 0;
    display:flex;align-items:center;gap:8px;}

/* ── ERROR / SUCCESS CARDS ── */
.error-card{background:#FDE8E8;border:1px solid #f0b8b8;border-left:4px solid #C0392B;
    border-radius:8px;padding:12px 16px;margin-bottom:8px;
    font-size:.83rem;color:#8B1A1A;display:flex;align-items:center;gap:10px;}
.success-card{background:#EAF6EE;border:1px solid #b8dfc4;border-left:4px solid #1E7A34;
    border-radius:8px;padding:13px 16px;margin:16px 60px 0 60px;
    font-size:.85rem;color:#1E5228;font-weight:600;display:flex;align-items:center;gap:10px;}

/* ── DIVIDER ── */
.divider{border:none;height:1px;background:#e8edf5;margin:28px 0;}

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
#  SESSION STATE INIT
# ─────────────────────────────────────────────
if "save_success" not in st.session_state:
    st.session_state.save_success = False


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
  <h2>Admin — Add New Player</h2>
  <p>Register a new player's profile and details into the master database</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  SECTION — PLAYER DETAILS FORM
# ─────────────────────────────────────────────
st.markdown(
    '<div class="sec-title"><span class="dot"></span>Player Details</div>',
    unsafe_allow_html=True
)
st.markdown('<div class="subsec-label">Identity and Profile</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    player_id   = st.number_input(
        "Player ID (Cap No)", min_value=100, help="Player ODI Cap No", key="player_id"
    )
    player_name = st.text_input("Player Name", key="player_name")
    country     = st.text_input("Country",value="India", key="country")
    role        = st.selectbox("Role", ["", "Batsman", "Bowler", "All-Rounder"], key="role")

with col2:
    bat_style  = st.selectbox(
        "Batting Style",
        ["", "RHB", "LHB", "NA"],
        key="bat_style"
    )
    bowl_style = st.selectbox(
        "Bowling Style",
        ["NA", "RA Fast", "LA Fast", "RA Medium", "LA Medium",
         "RA Off-Spin", "LA Off-Spin", "RA Leg-Break", "LA Leg-Break"],
        key="bowl_style"
    )
    today = date.today()

    dob = st.date_input(
    "Date of Birth",
    value=None,
    min_value=date(1970,1,1),
    max_value=today,
    key="dob"
    )

    debut_year = st.number_input(
        "Debut Year",
        placeholder=1990,
        max_value=date.today().year,
        step=1,
        key="debut_year"
        )
    
    debut_invalid = False
    if dob:
        min_debut_year = dob.year + 15
    
        if debut_year and debut_year < min_debut_year:
            st.error(f"Debut year must be >= {min_debut_year} (player must be at least 15).")
            debut_invalid = True
    image = st.file_uploader(
        "Upload Player Image",
        type=["jpg", "png", "jpeg"],
        key="image"
    )


# ─────────────────────────────────────────────
#  VALIDATION
# ─────────────────────────────────────────────
can_save     = True
is_duplicate = False

fields_incomplete = (
    player_id < 100         or
    not player_name.strip() or
    not country.strip()     or
    not role                or
    not bat_style           or
    image is None
)

if fields_incomplete:
    can_save = False

# Duplicate check (only when all fields filled)
if not fields_incomplete and os.path.exists(file_path):
    df_check = pd.read_csv(file_path)

    df_check["PlayerID"] = pd.to_numeric(df_check["PlayerID"], errors="coerce").fillna(0).astype(int)

    if int(player_id) in df_check["PlayerID"].values:
        can_save = False
        is_duplicate = True


# ─────────────────────────────────────────────
#  VALIDATION ERROR MESSAGE
# ─────────────────────────────────────────────
st.markdown('<hr class="divider">', unsafe_allow_html=True)

if is_duplicate:
    st.markdown(
        f'<div class="error-card">&nbsp;Player ID <strong>{int(player_id)}</strong> already exists — duplicate entry blocked.</div>',
        unsafe_allow_html=True
    )
elif fields_incomplete:
    st.markdown(
        '<div class="error-card">&nbsp;Please fill in all the required fields.</div>',
        unsafe_allow_html=True
    )


# ─────────────────────────────────────────────
#  ACTION BAR  —  Save button
# ─────────────────────────────────────────────
status_txt = "All fields complete — ready to save" if can_save else "Fill all required fields to enable Save"
status_cls = "save-status ok" if can_save else "save-status err"

st.markdown('<div class="action-bar">', unsafe_allow_html=True)
bar_l, bar_m, bar_r = st.columns([3, 2.5, 1.5])

with bar_l:
    st.markdown(f"""
    <div class="save-hint">
      <strong>{player_name.strip() or "—"}</strong>
      &nbsp;·&nbsp; {role or "No role selected"}<br>
      Cap No: <strong>{int(player_id) if player_id >= 100 else "—"}</strong>
      &nbsp;·&nbsp; {country.strip() or "—"}
      &nbsp;·&nbsp; Debut {int(debut_year)}
    </div>
    """, unsafe_allow_html=True)

with bar_m:
    st.markdown(
        f'<div class="{status_cls}" style="padding:8px 0;">{status_txt}</div>',
        unsafe_allow_html=True
    )

with bar_r:
    save = st.button("Save Player", disabled=not can_save, use_container_width=True)

st.markdown('</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  SUCCESS MESSAGE — rendered below the action bar
#  (flag set by save logic on previous run, shown
#   on the rerun, then cleared for the next run)
# ─────────────────────────────────────────────
if st.session_state.save_success:
    st.markdown(
        '<div class="success-card">&nbsp;Player saved successfully!</div>',
        unsafe_allow_html=True
    )
    st.session_state.save_success = False


# ─────────────────────────────────────────────
#  SAVE LOGIC
# ─────────────────────────────────────────────
if save:
    # Load existing data
    if os.path.exists(file_path):
        df_old = pd.read_csv(file_path)
        df_old["PlayerID"] = pd.to_numeric(df_old["PlayerID"], errors="coerce").fillna(0).astype(int)
    else:
        df_old = pd.DataFrame(columns=["PlayerID"])

    # Final duplicate guard
    if int(player_id) in df_old.get("PlayerID", pd.Series(dtype=int)).values:
        st.markdown(
            '<div class="error-card">&nbsp;Player ID already exists. Duplicate entry blocked.</div>',
            unsafe_allow_html=True
        )
        st.stop()

    # Save image
    img_name = f"{int(player_id)}_{player_name.replace(' ', '')}.jpg"
    with open(os.path.join(IMAGE_DIR, img_name), "wb") as f:
        f.write(image.getbuffer())

    new_row = {
        "PlayerID":      int(player_id),
        "PlayerName":    player_name,
        "Country":       country,
        "Role":          role,
        "BattingStyle":  bat_style,
        "BowlingStyle":  bowl_style,
        "DOB":           dob,
        "DebutYear":     int(debut_year),
        "ImageFileName": img_name,
        debut_invalid and "DebutInvalid" or "DebutValid": True
    }

    df_new = pd.DataFrame([new_row])

    if os.path.exists(file_path):
        df_old_raw = pd.read_csv(file_path)
        df_all     = pd.concat([df_old_raw, df_new], ignore_index=True)
        df_all.to_csv(file_path, index=False)
    else:
        df_new.to_csv(file_path, index=False)

    # Set success flag, clear all form keys, then rerun
    st.session_state.save_success = True
    for key in ["player_id", "player_name", "country", "role",
                "bat_style", "bowl_style", "dob", "debut_year", "image"]:
        if key in st.session_state:
            del st.session_state[key]

    st.rerun()


# ─────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div class="footer">
  Cricket Player Performance Analysis Platform &nbsp;|&nbsp;
  Built with Streamlit
</div>
""", unsafe_allow_html=True)