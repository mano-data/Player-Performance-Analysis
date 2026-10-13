import streamlit as st
import base64
import os

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Cricket Player Performance Analysis",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR  = os.path.join(BASE_DIR, "images")

def img_b64(filename):
    path = os.path.join(IMG_DIR, filename)
    if os.path.exists(path):
        ext  = filename.rsplit(".", 1)[-1].lower()
        mime = "jpeg" if ext in ["jpg", "jpeg"] else ext
        with open(path, "rb") as f:
            return f"data:image/{mime};base64,{base64.b64encode(f.read()).decode()}"
    return ""

stats_icon    = img_b64("stats_icon.png")
recent_icon   = img_b64("recent.png")
opponent_icon = img_b64("opponent.png")
venue_icon    = img_b64("venue.png")
strength_icon = img_b64("strength.png")
hero_img      = img_b64("cricket_ds.png")

# ─────────────────────────────────────────────
#  QUERY PARAM ROUTING
# ─────────────────────────────────────────────
query_params = st.query_params
if "page" in query_params:
    routes = {
        "overall":  "pages/Overall_performance.py",
        "recent":   "pages/recent_form.py",
        "opponent": "pages/Opponent_wise.py",
        "country":  "pages/Country_wise.py",
        "strength": "pages/Strength_weakness.py",
    }
    target = routes.get(query_params["page"])
    if target:
        st.switch_page(target)

# ─────────────────────────────────────────────
#  GLOBAL CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Poppins:wght@600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background: #ffffff !important;
}
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stSidebar"] { display: none; }

.block-container,
[data-testid="stAppViewBlockContainer"],
[data-testid="block-container"] {
    padding: 0 60px 60px 60px !important;
    max-width: 1560px !important;
    margin: 0 auto !important;
    background: #fff !important;
}
[data-testid="stAppViewBlockContainer"] > div:first-child { padding-top: 0 !important; }

/* ── NAVBAR ── */
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

/* ── HERO SECTION ── */
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
.hero-stats {
    display: flex;
    gap: 32px;
    margin-top: 8px;
}
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

/* ── SECTION DIVIDER ── */
.sec-divider {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, #1B3A6B 0%, #e8edf5 60%);
    margin: 8px 0 36px 0;
}
.sec-header {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 24px;
}
.sec-header-title {
    font-family: 'Poppins', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: #1B3A6B;
}
.sec-header-line {
    flex: 1;
    height: 1px;
    background: #e8edf5;
}

/* ── ANALYSIS CARDS ── */
.cards-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
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
.card-icon-emoji { font-size: 1.6rem; line-height: 1; }
.card-title {
    font-family: 'Poppins', sans-serif;
    font-size: 0.88rem;
    font-weight: 700;
    color: #1B3A6B;
    margin-bottom: 6px;
    text-decoration: none !important;
}
.analysis-card:hover .card-title { color: #1B3A6B; }
.card-desc { font-size: 0.72rem; color: #8899aa; line-height: 1.5; }

/* ── QUICK STATS STRIP ── */
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

/* ── HOW IT WORKS ── */
.how-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
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
.how-title {
    font-family: 'Poppins', sans-serif;
    font-weight: 700;
    font-size: 0.92rem;
    color: #1B3A6B;
    margin-bottom: 8px;
}
.how-desc { font-size: 0.8rem; color: #6B7A8D; line-height: 1.6; }

/* ── FOOTER ── */
.site-footer {
    background: #f7f9fc;
    border-top: 1px solid #dde3ee;
    padding: 24px 60px;
    text-align: center;
    margin: 0 -60px -60px -60px;
    color: #8899aa;
    font-size: 0.78rem;
    letter-spacing: 0.3px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  NAVBAR
# ─────────────────────────────────────────────
st.markdown("""
<div class="top-header">
  <div class="header-brand">
    Cricket Player<span> Performance Analysis</span>
  </div>
  <div class="header-nav">
    <a class="h-btn" href="/admin" target="_self">Admin</a>
    <a class="h-btn" href="/new" target="_self">Help</a>
  </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  HERO SECTION
# ─────────────────────────────────────────────
col_hero_l, col_hero_r = st.columns([1.3, 1], gap="large")

with col_hero_l:
    st.markdown("""
    <div style="padding: 48px 0 32px 0;">
        <div class="hero-eyebrow">Professional Cricket Analytics Platform</div>
        <div class="hero-title">
            Analyze Cricket<br>Performance
            <span> Like a Pro</span>
        </div>
        <p class="hero-desc">
            Deep insights into cricket player performance using structured data analytics.
            Evaluate overall stats, track recent form, analyze opponent &amp; country-wise
            performance, and identify strengths &amp; weaknesses — all in one place.
        </p>
        <div class="hero-stats">
            <div>
                <div class="hero-stat-val">10+</div>
                <div class="hero-stat-lbl">Players</div>
            </div>
            <div style="width:1px;background:#dde3ee;"></div>
            <div>
                <div class="hero-stat-val">5</div>
                <div class="hero-stat-lbl">Analysis Modules</div>
            </div>
            <div style="width:1px;background:#dde3ee;"></div>
            <div>
                <div class="hero-stat-val">100+</div>
                <div class="hero-stat-lbl">Match Records</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col_hero_r:
    st.markdown("<div style='padding-top:48px;'>", unsafe_allow_html=True)
    if hero_img:
        st.markdown(f"""
        <div style="border-radius:16px;overflow:hidden;border:1.5px solid #dde3ee;
                    box-shadow:0 8px 32px rgba(27,58,107,0.10);">
            <img src="{hero_img}" style="width:100%;display:block;">
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div style="background:#f7f9fc;border:1.5px solid #dde3ee;border-radius:16px;
                    height:280px;display:flex;align-items:center;justify-content:center;
                    font-size:4rem;">🏏</div>
        """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  QUICK STATS STRIP
# ─────────────────────────────────────────────
st.markdown("""
<div class="stats-strip">
    <div class="strip-item">
        <div class="strip-val">ODI</div>
        <div class="strip-lbl">Format</div>
    </div>
    <div class="strip-divider"></div>
    <div class="strip-item">
        <div class="strip-val">India</div>
        <div class="strip-lbl">Team</div>
    </div>
    <div class="strip-divider"></div>
    <div class="strip-item">
        <div class="strip-val">After 2019 CWC</div>
        <div class="strip-lbl">Data Period</div>
    </div>
    <div class="strip-divider"></div>
    <div class="strip-item">
        <div class="strip-val">Bat · Bowl · AR</div>
        <div class="strip-lbl">Player Roles</div>
    </div>
    <div class="strip-divider"></div>
    <div class="strip-item">
        <div class="strip-val">Real-time</div>
        <div class="strip-lbl">Analysis</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  ANALYSIS MODULE CARDS  — pure HTML, no st.button
#  Cards use ?page= query param so Streamlit's
#  query_params routing at the top handles navigation
# ─────────────────────────────────────────────
st.markdown("""
<div class="sec-header">
    <div class="sec-header-title">Analysis Modules</div>
    <div class="sec-header-line"></div>
</div>
""", unsafe_allow_html=True)

# Single, clean card definitions — no duplicate variable
card_definitions = [
    ("overall",  stats_icon,    "📊", "Overall Performance", "Batting &amp; bowling career stats at a glance"),
    ("recent",   recent_icon,   "📈", "Recent Form",         "Last 10–20 match performance trends"),
    ("opponent", opponent_icon, "🆚", "Opponent Wise",       "Head-to-head records vs each opponent"),
    ("country",  venue_icon,    "🌍", "Country Wise",        "Performance split by venue country"),
    ("strength", strength_icon, "💪", "Strength &amp; Weakness", "Identify key strengths and weak spots"),
]

def icon_html(b64_src, emoji_fallback):
    if b64_src:
        return f'<img src="{b64_src}" alt="">'
    return f'<span class="card-icon-emoji">{emoji_fallback}</span>'

card_html = '<div class="cards-grid">'
for page_key, b64, emoji, title, desc in card_definitions:
    card_html += f"""
    <a href="?page={page_key}" class="analysis-card" target="_self">
        <div class="card-icon-wrap">{icon_html(b64, emoji)}</div>
        <div class="card-title">{title}</div>
        <div class="card-desc">{desc}</div>
    </a>"""
card_html += "</div>"
st.markdown(card_html, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  HOW IT WORKS
# ─────────────────────────────────────────────
st.markdown("""
<div class="sec-divider"></div>
<div class="sec-header">
    <div class="sec-header-title">How It Works</div>
    <div class="sec-header-line"></div>
</div>
<div class="how-grid">
    <div class="how-card">
        <div class="how-step">1</div>
        <div class="how-title">Select a Player</div>
        <div class="how-desc">
            Choose any player from the dropdown — batsmen, bowlers, or all-rounders.
            The system auto-detects their role and adjusts the analysis accordingly.
        </div>
    </div>
    <div class="how-card">
        <div class="how-step">2</div>
        <div class="how-title">Choose an Analysis Module</div>
        <div class="how-desc">
            Pick from Overall Performance, Recent Form, Opponent Wise, Country Wise,
            or Strength &amp; Weakness depending on what insight you need.
        </div>
    </div>
    <div class="how-card">
        <div class="how-step">3</div>
        <div class="how-title">Explore &amp; Compare</div>
        <div class="how-desc">
            Dive into KPI cards, interactive charts, and data tables.
            Use the Compare module to benchmark two players side by side.
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<div class="site-footer">
    Cricket Player Performance Analysis &nbsp;|&nbsp; Built with Streamlit
</div>
""", unsafe_allow_html=True)