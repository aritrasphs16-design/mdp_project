"""
ChemoFilter — NOVA UI System v3.0
===================================
Complete visual reconstruction. Zero backend changes.
Modern SaaS-grade analytics dashboard design.

Design System: "Crystalline Obsidian"
- Deep navy/obsidian backgrounds
- Electric teal + amber dual-accent palette
- Sharp geometric aesthetics
- Card-based modular layout
- Premium typography: Syne + JetBrains Mono
"""

import streamlit as st
import assets


# ════════════════════════════════════════════════════════════════════════════
#  MAIN INJECTION — call once at app startup
# ════════════════════════════════════════════════════════════════════════════
def inject_ui():
    _inject_fonts()
    _inject_design_system()
    _inject_component_library()


def _inject_fonts():
    st.markdown("""
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=JetBrains+Mono:ital,wght@0,300;0,400;0,500;0,600;1,300&family=Inter:wght@300;400;500&display=swap" rel="stylesheet">
""", unsafe_allow_html=True)


def _inject_design_system():
    st.markdown("""
<style>
/* ════════════════════════════════════════════════════════════════════
   NOVA DESIGN TOKENS — Crystalline Obsidian
   ════════════════════════════════════════════════════════════════════ */
:root {
  /* Core surfaces */
  --n-bg:       #020408;
  --n-bg1:      #040a12;
  --n-bg2:      #060d18;
  --n-bg3:      #091220;
  --n-bg4:      #0d1a2e;
  --n-bg5:      #112038;

  /* Glass surfaces */
  --n-glass:    rgba(6, 20, 40, 0.6);
  --n-glass2:   rgba(10, 28, 54, 0.45);

  /* Borders */
  --n-bdr:      rgba(0, 210, 190, 0.12);
  --n-bdr2:     rgba(0, 210, 190, 0.07);
  --n-bdr3:     rgba(255, 255, 255, 0.05);
  --n-bdr4:     rgba(255, 255, 255, 0.025);

  /* Text */
  --n-tx:       #e8f4f0;
  --n-tx2:      rgba(200, 230, 220, 0.65);
  --n-tx3:      rgba(160, 200, 190, 0.38);
  --n-tx4:      rgba(120, 170, 160, 0.22);

  /* Primary accent: Electric Teal */
  --n-teal:     #00d2be;
  --n-teal2:    #00f5e0;
  --n-teal3:    #80e8df;
  --n-teal-bg:  rgba(0, 210, 190, 0.08);
  --n-teal-bdr: rgba(0, 210, 190, 0.25);
  --n-teal-glow: rgba(0, 210, 190, 0.15);

  /* Secondary accent: Amber */
  --n-amber:    #f0a020;
  --n-amber2:   #f5ba48;
  --n-amber3:   #fdd88a;
  --n-amber-bg: rgba(240, 160, 32, 0.08);
  --n-amber-bdr: rgba(240, 160, 32, 0.22);

  /* Status */
  --n-green:    #22d88a;
  --n-green-bg: rgba(34, 216, 138, 0.08);
  --n-red:      #ff5e6b;
  --n-red-bg:   rgba(255, 94, 107, 0.08);
  --n-yellow:   #f5c842;
  --n-violet:   #9b82f0;
  --n-blue:     #4aa3f5;

  /* Layout */
  --n-r:        14px;
  --n-r2:       10px;
  --n-r3:       6px;
  --n-shadow:   0 20px 60px rgba(0, 0, 0, 0.7), 0 4px 16px rgba(0, 0, 0, 0.4);
  --n-shadow2:  0 8px 32px rgba(0, 0, 0, 0.5);
  --n-shadow3:  0 2px 8px rgba(0, 0, 0, 0.35);

  /* Typography */
  --n-font-head: 'Syne', sans-serif;
  --n-font-body: 'Inter', sans-serif;
  --n-font-mono: 'JetBrains Mono', monospace;
}

/* ════════════ GLOBAL RESET ════════════ */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="block-container"],
.main, .block-container {
  background: var(--n-bg) !important;
  font-family: var(--n-font-body) !important;
  color: var(--n-tx) !important;
}

/* Grid overlay */
[data-testid="stAppViewContainer"]::before {
  content: '';
  position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background-image:
    linear-gradient(rgba(0,210,190,0.025) 1px, transparent 1px),
    linear-gradient(90deg, rgba(0,210,190,0.025) 1px, transparent 1px);
  background-size: 60px 60px;
}

/* Scanline effect */
[data-testid="stAppViewContainer"]::after {
  content: '';
  position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background: repeating-linear-gradient(
    0deg,
    transparent,
    transparent 3px,
    rgba(0,0,0,0.04) 3px,
    rgba(0,0,0,0.04) 4px
  );
}

/* Hide Streamlit chrome */
#MainMenu, footer, header,
[data-testid="stToolbar"],
.stDeployButton { visibility: hidden !important; }

/* ════════════ CONTAINER ════════════ */
[data-testid="block-container"] {
  padding: 2rem 3.5rem !important;
  max-width: 1480px !important;
  margin: 0 auto !important;
}

/* ════════════ SIDEBAR — Command Center ════════════ */
[data-testid="stSidebar"] {
  background: var(--n-bg1) !important;
  border-right: 1px solid var(--n-bdr) !important;
  backdrop-filter: blur(20px) !important;
}
[data-testid="stSidebar"]::before {
  content: '';
  position: absolute; top: 0; left: 0; right: 0; height: 3px;
  background: linear-gradient(90deg, var(--n-teal), var(--n-amber), var(--n-teal));
  background-size: 200%;
  animation: borderFlow 4s linear infinite;
}
@keyframes borderFlow {
  from { background-position: 0% 50%; }
  to { background-position: 200% 50%; }
}
[data-testid="stSidebar"] * { color: var(--n-tx) !important; }

/* Sidebar labels */
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stTextInput label,
[data-testid="stSidebar"] .stSlider label,
[data-testid="stSidebar"] .stRadio label {
  font-family: var(--n-font-mono) !important;
  font-size: 0.58rem !important;
  letter-spacing: 2px !important;
  text-transform: uppercase !important;
  color: var(--n-tx3) !important;
}

/* ════════════ TYPOGRAPHY ════════════ */
h1, h2, h3, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 {
  font-family: var(--n-font-head) !important;
  color: var(--n-tx) !important;
  letter-spacing: -0.5px !important;
}
p, .stMarkdown p, label, span {
  font-family: var(--n-font-body) !important;
  color: var(--n-tx2) !important;
}
code {
  font-family: var(--n-font-mono) !important;
  background: var(--n-bg3) !important;
  color: var(--n-teal) !important;
  border: 1px solid var(--n-bdr2) !important;
  border-radius: 4px !important;
  padding: 2px 6px !important;
  font-size: 0.82em !important;
}

/* ════════════ INPUTS ════════════ */
.stTextInput input, .stTextArea textarea {
  background: var(--n-bg3) !important;
  border: 1px solid var(--n-bdr) !important;
  border-radius: var(--n-r2) !important;
  color: var(--n-tx) !important;
  font-family: var(--n-font-mono) !important;
  font-size: 0.85rem !important;
  transition: all 0.2s ease !important;
  padding: 12px 16px !important;
}
.stTextInput input:focus, .stTextArea textarea:focus {
  border-color: var(--n-teal) !important;
  box-shadow: 0 0 0 3px var(--n-teal-bg), 0 0 20px var(--n-teal-glow) !important;
  outline: none !important;
}
.stTextInput input::placeholder, .stTextArea textarea::placeholder {
  color: var(--n-tx4) !important;
}

/* ════════════ BUTTONS ════════════ */
.stButton > button {
  font-family: var(--n-font-head) !important;
  font-weight: 600 !important;
  font-size: 0.78rem !important;
  letter-spacing: 1.5px !important;
  text-transform: uppercase !important;
  padding: 12px 28px !important;
  border-radius: var(--n-r2) !important;
  border: 1px solid var(--n-teal-bdr) !important;
  background: var(--n-teal-bg) !important;
  color: var(--n-teal) !important;
  cursor: pointer !important;
  transition: all 0.2s cubic-bezier(0.2, 0.8, 0.2, 1) !important;
  position: relative !important;
  overflow: hidden !important;
}
.stButton > button::before {
  content: '';
  position: absolute; inset: 0;
  background: linear-gradient(135deg, var(--n-teal), var(--n-teal2));
  opacity: 0;
  transition: opacity 0.2s;
}
.stButton > button:hover {
  border-color: var(--n-teal) !important;
  background: var(--n-teal-bg) !important;
  color: var(--n-teal2) !important;
  transform: translateY(-2px) !important;
  box-shadow: 0 8px 32px var(--n-teal-glow) !important;
}
.stButton > button:active {
  transform: translateY(0) !important;
}

/* Primary buttons (type="primary") */
.stButton > button[kind="primary"] {
  background: linear-gradient(135deg, var(--n-teal), var(--n-teal2)) !important;
  color: var(--n-bg) !important;
  border: none !important;
  font-weight: 700 !important;
}
.stButton > button[kind="primary"]:hover {
  box-shadow: 0 12px 40px var(--n-teal-glow) !important;
  color: var(--n-bg) !important;
}

/* Download buttons */
.stDownloadButton > button {
  font-family: var(--n-font-mono) !important;
  font-size: 0.65rem !important;
  letter-spacing: 1px !important;
  padding: 9px 18px !important;
  border-radius: var(--n-r3) !important;
  border: 1px solid var(--n-amber-bdr) !important;
  background: var(--n-amber-bg) !important;
  color: var(--n-amber) !important;
  transition: all 0.2s !important;
}
.stDownloadButton > button:hover {
  border-color: var(--n-amber) !important;
  box-shadow: 0 4px 20px rgba(240,160,32,0.2) !important;
  transform: translateY(-1px) !important;
}

/* ════════════ SELECTBOX ════════════ */
.stSelectbox > div > div {
  background: var(--n-bg3) !important;
  border: 1px solid var(--n-bdr) !important;
  border-radius: var(--n-r2) !important;
  color: var(--n-tx) !important;
  font-family: var(--n-font-mono) !important;
  font-size: 0.82rem !important;
}
.stSelectbox > div > div:focus-within {
  border-color: var(--n-teal) !important;
  box-shadow: 0 0 0 2px var(--n-teal-bg) !important;
}

/* ════════════ SLIDERS ════════════ */
.stSlider > div > div > div {
  background: var(--n-teal) !important;
}
.stSlider [data-baseweb="slider"] > div {
  background: var(--n-bg3) !important;
}
.stSlider [role="slider"] {
  background: var(--n-teal) !important;
  border: 2px solid var(--n-bg) !important;
  box-shadow: 0 0 10px var(--n-teal-glow) !important;
}

/* ════════════ TOGGLES / CHECKBOXES ════════════ */
.stCheckbox [data-baseweb="checkbox"] > div:first-child {
  background: var(--n-bg3) !important;
  border: 1px solid var(--n-bdr) !important;
  border-radius: 4px !important;
}
.stCheckbox [aria-checked="true"] > div:first-child {
  background: var(--n-teal) !important;
  border-color: var(--n-teal) !important;
}
.stToggle [data-baseweb="toggle"] > div {
  background: var(--n-bg4) !important;
  border: 1px solid var(--n-bdr) !important;
}
[data-checked="true"] .stToggle [data-baseweb="toggle"] > div {
  background: var(--n-teal) !important;
}

/* ════════════ TABS ════════════ */
.stTabs [data-baseweb="tab-list"] {
  background: var(--n-bg2) !important;
  border-bottom: 1px solid var(--n-bdr2) !important;
  gap: 2px !important;
  padding: 6px 8px 0 !important;
  border-radius: 12px 12px 0 0 !important;
  flex-wrap: wrap !important;
}
.stTabs [data-baseweb="tab"] {
  background: transparent !important;
  color: var(--n-tx3) !important;
  border: none !important;
  font-family: var(--n-font-mono) !important;
  font-size: 0.6rem !important;
  letter-spacing: 1px !important;
  text-transform: uppercase !important;
  padding: 8px 14px !important;
  border-radius: 6px 6px 0 0 !important;
  transition: all 0.2s !important;
}
.stTabs [aria-selected="true"] {
  background: var(--n-teal-bg) !important;
  color: var(--n-teal) !important;
  border-bottom: 2px solid var(--n-teal) !important;
}
.stTabs [data-baseweb="tab"]:hover:not([aria-selected="true"]) {
  color: var(--n-tx2) !important;
  background: var(--n-bg3) !important;
}
.stTabs [data-baseweb="tab-panel"] {
  background: var(--n-bg2) !important;
  border: 1px solid var(--n-bdr2) !important;
  border-top: none !important;
  border-radius: 0 0 12px 12px !important;
  padding: 20px !important;
}

/* ════════════ EXPANDERS ════════════ */
.stExpander {
  background: var(--n-bg2) !important;
  border: 1px solid var(--n-bdr2) !important;
  border-radius: var(--n-r) !important;
  margin-bottom: 8px !important;
  overflow: hidden !important;
  transition: border-color 0.2s !important;
}
.stExpander:hover { border-color: var(--n-bdr) !important; }
.stExpander summary {
  padding: 14px 20px !important;
  font-family: var(--n-font-mono) !important;
  font-size: 0.68rem !important;
  letter-spacing: 1.5px !important;
  text-transform: uppercase !important;
  color: var(--n-tx2) !important;
}
.stExpander [data-testid="stExpanderDetails"] {
  background: var(--n-bg2) !important;
  padding: 0 20px 16px !important;
}

/* ════════════ METRICS ════════════ */
[data-testid="stMetric"] {
  background: var(--n-bg2) !important;
  border: 1px solid var(--n-bdr2) !important;
  border-radius: var(--n-r) !important;
  padding: 16px 20px !important;
  transition: all 0.2s !important;
}
[data-testid="stMetric"]:hover {
  border-color: var(--n-bdr) !important;
  box-shadow: var(--n-shadow3) !important;
}
[data-testid="stMetric"] label {
  font-family: var(--n-font-mono) !important;
  font-size: 0.58rem !important;
  letter-spacing: 2px !important;
  text-transform: uppercase !important;
  color: var(--n-tx3) !important;
}
[data-testid="stMetricValue"] {
  font-family: var(--n-font-head) !important;
  font-size: 1.8rem !important;
  font-weight: 700 !important;
  color: var(--n-tx) !important;
}
[data-testid="stMetricDelta"] { font-family: var(--n-font-mono) !important; }

/* ════════════ DATAFRAMES ════════════ */
[data-testid="stDataFrame"] {
  border: 1px solid var(--n-bdr2) !important;
  border-radius: var(--n-r) !important;
  overflow: hidden !important;
}
[data-testid="stDataFrame"] thead th {
  background: var(--n-bg3) !important;
  color: var(--n-teal) !important;
  font-family: var(--n-font-mono) !important;
  font-size: 0.6rem !important;
  letter-spacing: 1.5px !important;
  text-transform: uppercase !important;
  border-bottom: 1px solid var(--n-bdr) !important;
  padding: 10px 12px !important;
}
[data-testid="stDataFrame"] tbody tr:nth-child(even) {
  background: rgba(0,210,190,0.025) !important;
}
[data-testid="stDataFrame"] tbody td {
  font-family: var(--n-font-mono) !important;
  font-size: 0.75rem !important;
  color: var(--n-tx2) !important;
  border-bottom: 1px solid var(--n-bdr4) !important;
  padding: 8px 12px !important;
}

/* ════════════ ALERTS / INFO ════════════ */
.stAlert {
  background: var(--n-bg2) !important;
  border-radius: var(--n-r2) !important;
  border: none !important;
}
[data-baseweb="notification"] {
  background: var(--n-bg2) !important;
  border-left: 3px solid var(--n-teal) !important;
  color: var(--n-tx2) !important;
  font-family: var(--n-font-mono) !important;
  font-size: 0.72rem !important;
  border-radius: 0 var(--n-r2) var(--n-r2) 0 !important;
}

/* ════════════ SPINNER ════════════ */
[data-testid="stSpinner"] {
  color: var(--n-teal) !important;
}

/* ════════════ SCROLLBARS ════════════ */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--n-bg1); }
::-webkit-scrollbar-thumb {
  background: var(--n-bdr);
  border-radius: 10px;
}
::-webkit-scrollbar-thumb:hover { background: var(--n-teal); }

/* ════════════ ANIMATIONS ════════════ */
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(8px); }
  to   { opacity: 1; transform: none; }
}
@keyframes pulseGlow {
  0%, 100% { box-shadow: 0 0 0 0 var(--n-teal-bg); }
  50%       { box-shadow: 0 0 0 8px transparent; }
}
@keyframes shimmerFlow {
  0%   { background-position: -200% center; }
  100% { background-position:  200% center; }
}
</style>
""", unsafe_allow_html=True)


def _inject_component_library():
    st.markdown("""
<style>
/* ════════════════════════════════════════════════════════════════════
   NOVA COMPONENT LIBRARY
   ════════════════════════════════════════════════════════════════════ */

/* ── HERO BANNER ── */
.hero {
  background: linear-gradient(135deg, var(--n-bg1) 0%, var(--n-bg3) 50%, var(--n-bg1) 100%);
  border: 1px solid var(--n-bdr);
  border-radius: 20px;
  padding: 48px 56px;
  margin-bottom: 28px;
  position: relative;
  overflow: hidden;
  animation: fadeIn 0.6s ease both;
}
.hero::before {
  content: '';
  position: absolute; top: 0; left: 0; right: 0; height: 1px;
  background: linear-gradient(90deg, transparent, var(--n-teal), var(--n-amber), var(--n-teal), transparent);
  background-size: 300%;
  animation: shimmerFlow 5s linear infinite;
}
.hero::after {
  content: '';
  position: absolute; bottom: 0; right: 0;
  width: 400px; height: 400px;
  background: radial-gradient(circle, rgba(0,210,190,0.04) 0%, transparent 70%);
  pointer-events: none;
}
.hero-hex {
  font-size: 1.5rem; color: var(--n-teal);
  display: inline-block; margin-bottom: 16px;
  text-shadow: 0 0 30px var(--n-teal-glow);
}
.hero-overline {
  font-family: var(--n-font-mono);
  font-size: 0.6rem; letter-spacing: 4px;
  text-transform: uppercase; color: var(--n-tx3);
  margin-bottom: 14px;
}
.hero-title {
  font-family: var(--n-font-head);
  font-size: 3.5rem; font-weight: 800;
  letter-spacing: -2px; line-height: 1;
  color: var(--n-tx); margin-bottom: 10px;
}
.hero-title span {
  background: linear-gradient(135deg, var(--n-teal), var(--n-teal2));
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text;
}
.hero-sub {
  font-family: var(--n-font-mono);
  font-size: 0.65rem; letter-spacing: 2px;
  color: var(--n-tx3); margin-bottom: 12px;
}
.hero-meta {
  font-size: 0.72rem; color: var(--n-tx3);
  font-family: var(--n-font-mono); line-height: 1.8;
}
.hero-stat-strip {
  position: absolute; top: 50%; right: 56px;
  transform: translateY(-50%); text-align: right;
}
.hss-num {
  font-family: var(--n-font-head);
  font-size: 4rem; font-weight: 800;
  color: var(--n-teal);
  text-shadow: 0 0 40px var(--n-teal-glow);
  line-height: 1;
}
.hss-lbl {
  font-family: var(--n-font-mono);
  font-size: 0.55rem; letter-spacing: 3px;
  color: var(--n-tx3); text-transform: uppercase;
}

/* ── FEATURE CHIPS ── */
.feature-chips { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 20px; }
.chip {
  font-family: var(--n-font-mono);
  font-size: 0.55rem; letter-spacing: 1.5px;
  text-transform: uppercase; padding: 5px 12px;
  border-radius: 30px; font-weight: 500;
}
.chip-teal {
  background: var(--n-teal-bg); color: var(--n-teal);
  border: 1px solid var(--n-teal-bdr);
}
.chip-amber {
  background: var(--n-amber-bg); color: var(--n-amber);
  border: 1px solid var(--n-amber-bdr);
}
.chip-gold { background: var(--n-teal-bg); color: var(--n-teal); border: 1px solid var(--n-teal-bdr); }
.chip-base { background: var(--n-bg3); color: var(--n-tx3); border: 1px solid var(--n-bdr3); }

/* ── STATS STRIP ── */
.stats-strip {
  display: grid;
  grid-template-columns: repeat(9, 1fr);
  gap: 1px;
  background: var(--n-bdr2);
  border: 1px solid var(--n-bdr);
  border-radius: var(--n-r);
  overflow: hidden;
  margin: 24px 0;
  animation: fadeIn 0.5s ease both;
}
.sc {
  background: var(--n-bg2);
  padding: 18px 12px;
  text-align: center;
  cursor: default;
  transition: background 0.2s;
}
.sc:hover { background: var(--n-bg3); }
.sc-val {
  font-family: var(--n-font-head);
  font-size: 1.6rem; font-weight: 700;
  line-height: 1; margin-bottom: 5px;
  cursor: help;
}
.sc-lbl {
  font-family: var(--n-font-mono);
  font-size: 0.46rem; letter-spacing: 2px;
  text-transform: uppercase; color: var(--n-tx3);
}

/* ── SECTION HEADERS ── */
.sec {
  display: flex; align-items: center; gap: 12px;
  margin: 32px 0 20px; padding-bottom: 14px;
  border-bottom: 1px solid var(--n-bdr2);
}
.sec-num {
  font-family: var(--n-font-mono);
  font-size: 0.6rem; color: var(--n-teal);
  opacity: 0.6; letter-spacing: 1px;
  background: var(--n-teal-bg);
  border: 1px solid var(--n-teal-bdr);
  padding: 3px 8px; border-radius: 4px;
}
.sec-title {
  font-family: var(--n-font-head);
  font-size: 1rem; font-weight: 700;
  color: var(--n-tx); letter-spacing: -0.3px;
}
.sec-line { flex: 1; height: 1px; background: var(--n-bdr2); }
.sec-tag {
  font-family: var(--n-font-mono);
  font-size: 0.54rem; letter-spacing: 1px;
  color: var(--n-tx4);
}

/* ── CARDS ── */
.card {
  background: var(--n-bg2);
  border: 1px solid var(--n-bdr2);
  border-radius: var(--n-r);
  padding: 20px;
  transition: border-color 0.2s, transform 0.2s;
  position: relative; overflow: hidden;
}
.card:hover { border-color: var(--n-bdr); }
.card-inner { position: relative; z-index: 1; }

/* ── AI PANEL ── */
.ai-panel {
  background: var(--n-bg2);
  border: 1px solid var(--n-bdr2);
  border-radius: var(--n-r);
  padding: 18px 20px;
  margin-bottom: 12px;
}
.ai-head {
  font-family: var(--n-font-mono);
  font-size: 0.58rem; letter-spacing: 2.5px;
  text-transform: uppercase; color: var(--n-teal);
  opacity: 0.7; margin-bottom: 14px;
  padding-bottom: 10px;
  border-bottom: 1px solid var(--n-bdr2);
}
.ai-body {
  font-size: 0.82rem; line-height: 1.75;
  color: var(--n-tx2); font-family: var(--n-font-body);
}

/* ── REF BOX ── */
.ref-box {
  background: var(--n-bg2);
  border: 1px solid var(--n-bdr2);
  border-radius: var(--n-r);
  padding: 18px 24px;
  margin-bottom: 20px;
  position: relative; overflow: hidden;
}
.ref-box::before {
  content: '';
  position: absolute; left: 0; top: 0; bottom: 0; width: 3px;
  background: linear-gradient(180deg, var(--n-teal), var(--n-amber));
}
.ref-name {
  font-family: var(--n-font-head);
  font-size: 0.95rem; font-weight: 700;
  color: var(--n-tx); margin-bottom: 4px;
}

/* ── DESCRIPTOR TABLE ── */
.dtable {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.74rem;
  font-family: var(--n-font-mono);
  margin-bottom: 14px;
}
.dtable td { padding: 7px 10px; }
.dk {
  color: var(--n-tx3);
  border-bottom: 1px solid var(--n-bdr4);
  letter-spacing: 0.5px;
}
.dv {
  text-align: right;
  color: var(--n-tx);
  font-weight: 500;
  border-bottom: 1px solid var(--n-bdr4);
}
.dv.ok  { color: var(--n-green); }
.dv.warn { color: var(--n-yellow); }
.dv.bad  { color: var(--n-red); }

/* ── PROGRESS BARS ── */
.bar-lbl {
  font-family: var(--n-font-mono);
  font-size: 0.58rem; color: var(--n-tx3);
  letter-spacing: 1px; margin-bottom: 4px;
  text-transform: uppercase;
}
.bar-track {
  height: 3px;
  background: var(--n-bg4);
  border-radius: 10px; overflow: hidden;
  margin-bottom: 3px;
}
.bar-fill {
  height: 100%;
  border-radius: 10px;
  transition: width 0.6s cubic-bezier(0.2, 0.8, 0.2, 1);
}
.bar-num {
  font-family: var(--n-font-mono);
  font-size: 0.56rem; color: var(--n-tx4);
  text-align: right;
}

/* ── TOX PILLS ── */
.tpill {
  display: flex; align-items: center; gap: 8px;
  padding: 8px 12px; border-radius: var(--n-r2);
  font-size: 0.72rem; font-family: var(--n-font-mono);
  margin-bottom: 6px;
  border: 1px solid transparent;
  letter-spacing: 0.3px;
}
.tp-ok   { background: var(--n-green-bg); color: var(--n-green); border-color: rgba(34,216,138,0.2); }
.tp-warn { background: rgba(245,200,66,0.06); color: var(--n-yellow); border-color: rgba(245,200,66,0.2); }
.tp-bad  { background: var(--n-red-bg); color: var(--n-red); border-color: rgba(255,94,107,0.2); }

/* ── OPTIMISATION BOX ── */
.opt-box {
  background: var(--n-bg3);
  border: 1px solid var(--n-bdr2);
  border-radius: var(--n-r);
  padding: 14px 16px;
  margin: 12px 0;
}
.opt-head {
  font-family: var(--n-font-mono);
  font-size: 0.56rem; letter-spacing: 2px;
  text-transform: uppercase; color: var(--n-amber);
  margin-bottom: 10px; padding-bottom: 8px;
  border-bottom: 1px solid var(--n-bdr4);
}
.opt-row {
  display: flex; gap: 10px; padding: 4px 0;
  border-bottom: 1px solid var(--n-bdr4);
  font-family: var(--n-font-mono);
}
.opt-k {
  font-size: 0.64rem; color: var(--n-teal);
  min-width: 110px; font-weight: 500;
}
.opt-v { font-size: 0.62rem; color: var(--n-tx3); }

/* ── VERDICT BANNERS ── */
.verdict {
  border-radius: var(--n-r2);
  padding: 12px 16px; margin-top: 12px;
  border: 1px solid;
}
.vgo   { background: var(--n-green-bg); border-color: rgba(34,216,138,0.25); }
.vwarn { background: rgba(245,200,66,0.05); border-color: rgba(245,200,66,0.2); }
.vstop { background: var(--n-red-bg); border-color: rgba(255,94,107,0.2); }
.vt {
  font-family: var(--n-font-head);
  font-size: 0.82rem; font-weight: 700;
  margin-bottom: 3px;
}
.vgo .vt   { color: var(--n-green); }
.vwarn .vt { color: var(--n-yellow); }
.vstop .vt { color: var(--n-red); }
.vb {
  font-family: var(--n-font-mono);
  font-size: 0.6rem; color: var(--n-tx3);
  letter-spacing: 0.5px;
}

/* ── ANALYSIS CARDS ── */
.ana-card {
  background: var(--n-bg3);
  border: 1px solid var(--n-bdr2);
  border-radius: var(--n-r2);
  padding: 10px 14px;
  margin-bottom: 6px;
  transition: border-color 0.2s;
}
.ana-card:hover { border-color: var(--n-bdr); }
.ana-n {
  font-family: var(--n-font-mono);
  font-size: 0.72rem; color: var(--n-tx);
  font-weight: 500; margin-bottom: 3px;
}
.ana-ex {
  font-size: 0.65rem; color: var(--n-tx3);
  font-family: var(--n-font-body);
}

/* ── MEDALLION GRADES ── */
.medallion-wrap {
  display: flex; align-items: center; gap: 10px;
  margin-bottom: 12px;
}
.medallion {
  width: 38px; height: 38px; border-radius: 10px;
  display: flex; align-items: center; justify-content: center;
  font-family: var(--n-font-head); font-size: 1.1rem;
  font-weight: 800;
}
.mA { background: rgba(34,216,138,0.12); color: var(--n-green); border: 1.5px solid rgba(34,216,138,0.3); }
.mB { background: var(--n-teal-bg); color: var(--n-teal); border: 1.5px solid var(--n-teal-bdr); }
.mC { background: rgba(245,200,66,0.08); color: var(--n-yellow); border: 1.5px solid rgba(245,200,66,0.2); }
.mF { background: var(--n-red-bg); color: var(--n-red); border: 1.5px solid rgba(255,94,107,0.25); }
.med-id {
  font-family: var(--n-font-head);
  font-size: 0.88rem; font-weight: 700;
  color: var(--n-tx);
}

/* ── ROW DISPLAY ── */
.rrow {
  display: flex; align-items: baseline;
  justify-content: space-between;
  padding: 5px 0;
  border-bottom: 1px solid var(--n-bdr4);
  font-family: var(--n-font-mono);
}
.rk { font-size: 0.62rem; color: var(--n-tx3); letter-spacing: 0.5px; }
.rv {
  font-size: 0.7rem; color: var(--n-tx);
  font-weight: 500; text-align: right;
}
.rh {
  font-size: 0.54rem; letter-spacing: 2px;
  text-transform: uppercase; color: var(--n-teal);
  opacity: 0.6; margin: 10px 0 4px;
  font-family: var(--n-font-mono);
}

/* ── TAGS & PILLS ── */
.tag {
  font-family: var(--n-font-mono);
  font-size: 0.55rem; letter-spacing: 1px;
  padding: 3px 8px; border-radius: 20px;
  display: inline-block; margin: 2px;
}
.tag-a { background: var(--n-teal-bg); color: var(--n-teal); border: 1px solid var(--n-teal-bdr); }
.tag-b { background: var(--n-bg4); color: var(--n-tx3); border: 1px solid var(--n-bdr3); }
.tag-c { background: var(--n-amber-bg); color: var(--n-amber); border: 1px solid var(--n-amber-bdr); }

/* ── META SITE CARDS ── */
.meta-site {
  background: var(--n-bg3);
  border: 1px solid var(--n-bdr2);
  border-left: 3px solid var(--n-teal);
  border-radius: 0 var(--n-r2) var(--n-r2) 0;
  padding: 10px 14px;
  margin: 6px 0;
  font-family: var(--n-font-mono);
  transition: border-color 0.2s;
}
.ms-high { border-left-color: var(--n-red); }

/* ── REPORT BLOCK ── */
.rblock {
  background: var(--n-bg2);
  border: 1px solid var(--n-bdr2);
  border-radius: var(--n-r);
  padding: 20px 24px;
  margin-bottom: 12px;
  animation: fadeIn 0.3s ease both;
}

/* ── AURA IMAGE ── */
.aura-img {
  filter: brightness(0.95) contrast(1.05) saturate(0.9);
  transition: all 0.3s;
}
.aura-img:hover {
  filter: brightness(1) contrast(1.1) saturate(1.1);
}
.pulse-img { animation: pulseGlow 4s ease-in-out infinite; }

/* ── SIDEBAR BRAND ── */
.sidebar-brand {
  padding: 20px 16px 16px;
  border-bottom: 1px solid var(--n-bdr2);
  margin-bottom: 12px;
}
.sb-logo {
  font-family: var(--n-font-head);
  font-size: 1.4rem; font-weight: 800;
  color: var(--n-tx); letter-spacing: -0.5px;
  display: flex; align-items: center; gap: 8px;
}
.sb-hex {
  width: 32px; height: 32px; border-radius: 8px;
  background: linear-gradient(135deg, var(--n-teal), var(--n-teal2));
  display: flex; align-items: center; justify-content: center;
  font-size: 1rem; color: var(--n-bg);
  box-shadow: 0 4px 16px var(--n-teal-glow);
}
.sb-badge {
  font-family: var(--n-font-mono);
  font-size: 0.45rem; letter-spacing: 2px;
  text-transform: uppercase; color: var(--n-tx3);
  margin-top: 3px;
}

/* ── FOOTER ── */
.footer {
  text-align: center;
  padding: 40px 20px;
  margin-top: 60px;
  border-top: 1px solid var(--n-bdr2);
  font-family: var(--n-font-mono);
  font-size: 0.52rem; letter-spacing: 2px;
  color: var(--n-tx4); line-height: 2.5;
}

/* ── FILTER RESULTS TABLE ── */
.filter-table {
  width: 100%;
  border-collapse: collapse;
  font-family: var(--n-font-mono);
  font-size: 0.68rem;
}
.filter-table th {
  background: var(--n-bg3);
  color: var(--n-teal);
  padding: 8px 12px;
  text-align: left;
  font-size: 0.56rem;
  letter-spacing: 1.5px;
  text-transform: uppercase;
  border-bottom: 1px solid var(--n-bdr);
}
.filter-table td {
  padding: 7px 12px;
  border-bottom: 1px solid var(--n-bdr4);
  color: var(--n-tx2);
}
.filter-table tr:hover td { background: rgba(0,210,190,0.03); }

/* ── BATCH INTEL SECTION ── */
.batch-stat {
  background: var(--n-bg3);
  border: 1px solid var(--n-bdr2);
  border-radius: var(--n-r2);
  padding: 14px;
  text-align: center;
}
.batch-stat-n {
  font-family: var(--n-font-head);
  font-size: 1.6rem; font-weight: 700;
  color: var(--n-teal); line-height: 1;
}
.batch-stat-l {
  font-family: var(--n-font-mono);
  font-size: 0.52rem; color: var(--n-tx3);
  letter-spacing: 1.5px; text-transform: uppercase;
  margin-top: 4px;
}

/* ── SCORE BADGE ── */
.score-badge {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 4px 10px; border-radius: 20px;
  font-family: var(--n-font-mono);
  font-size: 0.6rem; font-weight: 600;
}

/* ── INFO PANEL ── */
.info-panel {
  display: flex; gap: 8px; padding: 12px;
  border-radius: var(--n-r2); margin-bottom: 10px;
  font-size: 0.74rem; font-family: var(--n-font-body);
  background: var(--n-bg3); border: 1px solid var(--n-bdr2);
  color: var(--n-tx2);
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
#  COMPONENT FUNCTIONS  (signatures must match app.py imports exactly)
# ════════════════════════════════════════════════════════════════════════════

def render_section_header(number: str, title: str, tag: str = ""):
    st.markdown(f"""
<div class="sec">
  <span class="sec-num">{number}</span>
  <span class="sec-title">{title}</span>
  <div class="sec-line"></div>
  {"<span class='sec-tag'>" + tag + "</span>" if tag else ""}
</div>
""", unsafe_allow_html=True)


def render_compound_header(cid: str, grade: str, score: int):
    color = {"A": "#22d88a", "B": "#00d2be", "C": "#f5c842", "F": "#ff5e6b"}.get(grade, "#aaa")
    st.markdown(f"""
<div style="display:flex;align-items:center;gap:14px;margin-bottom:16px">
  <div style="width:42px;height:42px;border-radius:11px;background:rgba(0,210,190,0.08);
    border:1.5px solid rgba(0,210,190,0.25);display:flex;align-items:center;justify-content:center;
    font-family:var(--n-font-head);font-size:1.2rem;font-weight:800;color:{color}">{grade}</div>
  <div>
    <div style="font-family:var(--n-font-head);font-size:1rem;font-weight:700;color:var(--n-tx)">{cid}</div>
    <div style="font-family:var(--n-font-mono);font-size:0.56rem;color:var(--n-tx3);letter-spacing:1px">
      LEAD SCORE {score}/100
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


def render_score_badge(score: int, label: str = ""):
    if score >= 75:
        color, bg, bdr = "#22d88a", "rgba(34,216,138,0.1)", "rgba(34,216,138,0.25)"
    elif score >= 50:
        color, bg, bdr = "#00d2be", "rgba(0,210,190,0.1)", "rgba(0,210,190,0.25)"
    elif score >= 25:
        color, bg, bdr = "#f5c842", "rgba(245,200,66,0.1)", "rgba(245,200,66,0.25)"
    else:
        color, bg, bdr = "#ff5e6b", "rgba(255,94,107,0.1)", "rgba(255,94,107,0.25)"
    st.markdown(f"""
<span class="score-badge" style="background:{bg};color:{color};border:1px solid {bdr}">
  {label or "SCORE"}&nbsp;{score}
</span>
""", unsafe_allow_html=True)


def render_metric_card(label: str, value, color: str = "var(--n-teal)", tooltip: str = ""):
    st.markdown(f"""
<div class="batch-stat" title="{tooltip}">
  <div class="batch-stat-n" style="color:{color}">{value}</div>
  <div class="batch-stat-l">{label}</div>
</div>
""", unsafe_allow_html=True)


def render_info_panel(text: str, kind: str = "info"):
    icons = {"info": "ℹ", "success": "✓", "warning": "⚠", "error": "✕"}
    colors = {"info": "var(--n-teal)", "success": "var(--n-green)", "warning": "var(--n-yellow)", "error": "var(--n-red)"}
    icon = icons.get(kind, "ℹ")
    color = colors.get(kind, "var(--n-teal)")
    st.markdown(f"""
<div class="info-panel">
  <span style="color:{color};font-size:0.9rem;flex-shrink:0">{icon}</span>
  <span>{text}</span>
</div>
""", unsafe_allow_html=True)


def render_progress_bar(label: str, value: float, max_val: float, color: str = "var(--n-teal)"):
    pct = min(100, value / max_val * 100) if max_val else 0
    st.markdown(f"""
<div style="margin-bottom:10px">
  <div class="bar-lbl">{label}</div>
  <div class="bar-track">
    <div class="bar-fill" style="width:{pct:.1f}%;background:{color}"></div>
  </div>
  <div class="bar-num">{value:.1f} / {max_val:.0f}</div>
</div>
""", unsafe_allow_html=True)


def render_pill(text: str, kind: str = "a"):
    cls_map = {"a": "tag-a", "b": "tag-b", "c": "tag-c"}
    st.markdown(f'<span class="tag {cls_map.get(kind, "tag-b")}">{text}</span>', unsafe_allow_html=True)


def render_tox_alert(level: str, label: str, detail: str = ""):
    cls = {"LOW": "tp-ok", "MEDIUM": "tp-warn", "HIGH": "tp-bad"}.get(level, "tp-warn")
    icon = {"LOW": "✓", "MEDIUM": "⚠", "HIGH": "✕"}.get(level, "—")
    st.markdown(f"""
<div class="tpill {cls}">
  <span style="font-size:0.85rem">{icon}</span>
  <span><b>{label}: {level}</b>{(" — " + detail) if detail else ""}</span>
</div>
""", unsafe_allow_html=True)


def render_ai_response(text: str):
    st.markdown(f"""
<div class="ai-panel">
  <div class="ai-head">⬡ AI ANALYSIS — CLAUDE</div>
  <div class="ai-body">{text}</div>
</div>
""", unsafe_allow_html=True)


def render_filter_results_table(tests: list):
    if not tests:
        return
    cats = {}
    for t in tests:
        cats.setdefault(t.get("category", "Other"), []).append(t)
    for cat, items in cats.items():
        st.markdown(f"""
<div style="margin-bottom:16px">
  <div style="font-family:var(--n-font-mono);font-size:0.55rem;letter-spacing:2px;
    text-transform:uppercase;color:var(--n-teal);opacity:0.6;margin-bottom:8px">{cat}</div>
  <table class="filter-table">
    <thead><tr><th>Test</th><th>Result</th><th>Detail</th></tr></thead>
    <tbody>
""", unsafe_allow_html=True)
        for item in items:
            res = item.get("result", "INFO")
            color = {"PASS": "#22d88a", "FAIL": "#ff5e6b", "INFO": "#00d2be", "WARN": "#f5c842"}.get(res, "#aaa")
            st.markdown(f"""
      <tr>
        <td>{item.get("test", "")}</td>
        <td style="color:{color};font-weight:600">{res}</td>
        <td>{item.get("detail", "")}</td>
      </tr>
""", unsafe_allow_html=True)
        st.markdown("</tbody></table></div>", unsafe_allow_html=True)


def render_sidebar_brand():
    st.sidebar.markdown("""
<div class="sidebar-brand">
  <div class="sb-logo">
    <div class="sb-hex">⬡</div>
    ChemoFilter
  </div>
  <div class="sb-badge">ADMET · Drug Discovery · VIT 2026</div>
</div>
""", unsafe_allow_html=True)


def theme_toggle_sidebar():
    # Theme toggle kept for compatibility — no-op in new design
    pass
