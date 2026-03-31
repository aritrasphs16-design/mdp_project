"""
ChemoFilter — NOVA Landing Page v3.0
=====================================
Complete visual reconstruction.
Crystalline Obsidian aesthetic with electric teal accents.
"""
import streamlit as st


def render_landing() -> bool:
    """Render the landing page. Returns True when user clicks enter."""
    st.markdown(_LANDING_CSS, unsafe_allow_html=True)
    st.markdown(_LANDING_HTML, unsafe_allow_html=True)

    # Centered launch button via Streamlit
    col1, col2, col3 = st.columns([1.8, 1, 1.8])
    with col2:
        launched = st.button("⬡  Launch ChemoFilter", key="_lp_launch", use_container_width=True)

    st.markdown(_POST_BUTTON_HTML, unsafe_allow_html=True)

    col4, col5, col6 = st.columns([2, 1, 2])
    with col5:
        cta = st.button("Begin Discovery →", key="_lp_cta", use_container_width=True)

    st.markdown(_LOWER_SECTIONS, unsafe_allow_html=True)

    # ── ADDITIVE ENHANCEMENTS (do not modify anything above) ──────────────────
    # All 10 interactive enhancements: molecule preview, sample selector,
    # validation, insight panel, feature cards, memory, advanced toggle,
    # API health, tooltips, micro-animations.
    try:
        from landing_enhancements import render_landing_enhancements
        render_landing_enhancements()
    except Exception:
        pass  # enhancements are purely additive — never block launch

    return launched or cta


_LANDING_CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;500;600;700;800&family=JetBrains+Mono:wght@300;400;500&family=Inter:wght@300;400;500&display=swap');

:root {
  --l-bg:    #020408;
  --l-bg1:   #040a12;
  --l-bg2:   #06101e;
  --l-bdr:   rgba(0,210,190,0.12);
  --l-bdr2:  rgba(0,210,190,0.06);
  --l-bdr3:  rgba(255,255,255,0.05);
  --l-tx:    #e4eeec;
  --l-tx2:   rgba(180,220,215,0.6);
  --l-tx3:   rgba(130,180,170,0.35);
  --l-teal:  #00d2be;
  --l-teal2: #00f5e0;
  --l-amber: #f0a020;
  --l-green: #22d88a;
  --l-glow:  rgba(0,210,190,0.14);
}

/* Reset Streamlit for landing */
#MainMenu, footer, header,
[data-testid="stToolbar"],
.stDeployButton { visibility: hidden !important; }
section[data-testid="stSidebar"] { display: none !important; }

html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="block-container"],
.main, .block-container {
  background: var(--l-bg) !important;
  padding: 0 !important;
  max-width: 100% !important;
  margin: 0 !important;
}
[data-testid="stVerticalBlock"] { gap: 0 !important; }

/* Launch button style */
.stButton > button {
  font-family: 'Syne', sans-serif !important;
  font-weight: 700 !important;
  font-size: 1rem !important;
  letter-spacing: 1px !important;
  padding: 16px 40px !important;
  border-radius: 12px !important;
  background: linear-gradient(135deg, var(--l-teal), var(--l-teal2)) !important;
  color: #020408 !important;
  border: none !important;
  cursor: pointer !important;
  transition: all 0.3s cubic-bezier(0.2, 0.8, 0.2, 1) !important;
  box-shadow: 0 8px 40px var(--l-glow) !important;
  width: 100% !important;
}
.stButton > button:hover {
  transform: translateY(-3px) !important;
  box-shadow: 0 16px 60px rgba(0,210,190,0.4), 0 0 0 6px rgba(0,210,190,0.08) !important;
}

/* Animations */
@keyframes lFadeUp {
  from { opacity: 0; transform: translateY(20px); }
  to   { opacity: 1; transform: none; }
}
@keyframes lOrb1 {
  0%   { transform: translate(0, 0) scale(1); }
  100% { transform: translate(60px, 40px) scale(1.2); }
}
@keyframes lOrb2 {
  0%   { transform: translate(0, 0) scale(1.1); }
  100% { transform: translate(-50px, -60px) scale(1); }
}
@keyframes lOrb3 {
  0%, 100% { transform: translate(0, 0); }
  50%      { transform: translate(30px, -20px); }
}
@keyframes lFlow {
  0%   { background-position: 0 50%; }
  100% { background-position: 300% 50%; }
}
@keyframes lHexSpin {
  0%   { transform: rotate(0deg) scale(1); }
  50%  { transform: rotate(3deg) scale(1.05); }
  100% { transform: rotate(0deg) scale(1); }
}
@keyframes lTickerScroll {
  from { transform: translateX(0); }
  to   { transform: translateX(-50%); }
}
@keyframes lPulse {
  0%, 100% { opacity: 0.6; }
  50%       { opacity: 1; }
}
@keyframes lScan {
  0%   { top: -100%; }
  100% { top: 200%; }
}
</style>"""


_LANDING_HTML = """
<div id="cf-landing">

<!-- ── AMBIENT ORBS ── -->
<div class="l-orbs">
  <div class="l-o l-o1"></div>
  <div class="l-o l-o2"></div>
  <div class="l-o l-o3"></div>
  <div class="l-o l-o4"></div>
  <div class="l-o l-o5"></div>
</div>

<!-- ── GRID OVERLAY ── -->
<div class="l-grid"></div>

<!-- ── SCAN LINE ── -->
<div class="l-scan"></div>

<!-- ── NAVIGATION ── -->
<nav class="l-nav">
  <div class="l-brand">
    <div class="l-hexbadge">⬡</div>
    <span>ChemoFilter</span>
  </div>
  <div class="l-navlinks">
    <span class="l-nl">ADMET Analysis</span>
    <span class="l-nl">Drug Discovery</span>
    <span class="l-nl">Visualization</span>
    <span class="l-nl">AI Engine</span>
  </div>
  <div class="l-ver">v1,000,000 · OMNIPOTENT</div>
</nav>

<!-- ── HERO SECTION ── -->
<section class="l-hero">

  <div class="l-eyebrow">
    <div class="l-edot"></div>
    Computational Drug Discovery Platform · VIT Chennai MDP 2026
  </div>

  <h1 class="l-headline">
    <span class="l-hl1">Chemo</span><em>Filter</em>
  </h1>
  <h2 class="l-subhead">Crystalline Omnipotence Edition</h2>

  <p class="l-desc">
    21-parameter ADMET intelligence powered by <strong>RDKit</strong>,
    <strong>Claude AI</strong>, and the world's most sophisticated
    molecular scoring engines. From Lipinski to Aether v10000 —
    every compound, every dimension.
  </p>

  <div class="l-chips">
    <span class="l-chip l-chip-t">Lipinski Ro5</span>
    <span class="l-chip l-chip-t">BOILED-EGG</span>
    <span class="l-chip l-chip-t">QED Scoring</span>
    <span class="l-chip l-chip-t">SA Score</span>
    <span class="l-chip l-chip-t">CYP Panel ×5</span>
    <span class="l-chip l-chip-t">hERG Risk</span>
    <span class="l-chip l-chip-t">CNS MPO</span>
    <span class="l-chip l-chip-t">PAINS Filter</span>
    <span class="l-chip l-chip-a">Lead Score™</span>
    <span class="l-chip l-chip-a">Oral Bio Score</span>
    <span class="l-chip l-chip-a">AI Explainer</span>
    <span class="l-chip l-chip-a">Drug Repurposing</span>
    <span class="l-chip l-chip-d">Aether v10000</span>
    <span class="l-chip l-chip-d">Xenon v5000</span>
    <span class="l-chip l-chip-d">Celestial v1000</span>
    <span class="l-chip l-chip-d">Omega v2000</span>
  </div>

</section>

<!-- ── STAT BAR ── -->
<div class="l-statbar">
  <div class="l-stat">
    <div class="l-sv" style="color:var(--l-teal)">21<span>+</span></div>
    <div class="l-sl">ADMET Parameters</div>
  </div>
  <div class="l-stat">
    <div class="l-sv" style="color:var(--l-amber)">200<span>+</span></div>
    <div class="l-sl">FDA Drug Anchors</div>
  </div>
  <div class="l-stat">
    <div class="l-sv" style="color:var(--l-green)">100k<span>+</span></div>
    <div class="l-sl">Feature Tensors</div>
  </div>
  <div class="l-stat">
    <div class="l-sv" style="color:#a78bfa">202</div>
    <div class="l-sl">Scientific Refs</div>
  </div>
  <div class="l-stat">
    <div class="l-sv" style="color:var(--l-teal)">9</div>
    <div class="l-sl">Engine Tiers</div>
  </div>
  <div class="l-stat">
    <div class="l-sv" style="color:var(--l-amber)">99.9<span>%</span></div>
    <div class="l-sl">Clinical Confidence</div>
  </div>
</div>

<!-- ── TICKER ── -->
<div class="l-ticker-wrap">
  <div class="l-ticker">
    <div class="l-ticker-inner">
      <span>RDKit</span><span class="l-tdot">·</span>
      <span>Lipinski Rule of Five</span><span class="l-tdot">·</span>
      <span>BOILED-EGG</span><span class="l-tdot">·</span>
      <span>QED Scoring</span><span class="l-tdot">·</span>
      <span>SA Score</span><span class="l-tdot">·</span>
      <span>Tanimoto Similarity</span><span class="l-tdot">·</span>
      <span>CYP Inhibition Panel</span><span class="l-tdot">·</span>
      <span>hERG Risk</span><span class="l-tdot">·</span>
      <span>CNS MPO</span><span class="l-tdot">·</span>
      <span>PAINS Filter</span><span class="l-tdot">·</span>
      <span>ECFP4 Fingerprints</span><span class="l-tdot">·</span>
      <span>ESOL Solubility</span><span class="l-tdot">·</span>
      <span>Murcko Scaffolds</span><span class="l-tdot">·</span>
      <span>Veber Rule</span><span class="l-tdot">·</span>
      <span>Egan Filter</span><span class="l-tdot">·</span>
      <span>Ghose Filter</span><span class="l-tdot">·</span>
      <span>Lead Score™</span><span class="l-tdot">·</span>
      <span>Oral Bio Score</span><span class="l-tdot">·</span>
      <span>Promiscuity Risk</span><span class="l-tdot">·</span>
      <span>Metabolic Pulse</span><span class="l-tdot">·</span>
      <span>3D Conformers</span><span class="l-tdot">·</span>
      <span>Claude AI Explainer</span><span class="l-tdot">·</span>
      <!-- Repeat for seamless loop -->
      <span>RDKit</span><span class="l-tdot">·</span>
      <span>Lipinski Rule of Five</span><span class="l-tdot">·</span>
      <span>BOILED-EGG</span><span class="l-tdot">·</span>
      <span>QED Scoring</span><span class="l-tdot">·</span>
      <span>SA Score</span><span class="l-tdot">·</span>
      <span>Tanimoto Similarity</span><span class="l-tdot">·</span>
      <span>CYP Inhibition Panel</span><span class="l-tdot">·</span>
      <span>hERG Risk</span><span class="l-tdot">·</span>
      <span>CNS MPO</span><span class="l-tdot">·</span>
    </div>
  </div>
</div>

<style>
#cf-landing {
  font-family: 'Inter', sans-serif;
  color: var(--l-tx);
  background: var(--l-bg);
  position: relative;
  overflow-x: hidden;
  min-height: 100vh;
}
.l-orbs { position: fixed; inset: 0; pointer-events: none; z-index: 0; overflow: hidden; }
.l-o {
  position: absolute; border-radius: 50%;
  filter: blur(120px); opacity: 0.12;
}
.l-o1 {
  width: 1100px; height: 1100px;
  top: -400px; right: -300px;
  background: radial-gradient(circle, #00d2be 0%, transparent 65%);
  animation: lOrb1 20s ease-in-out infinite alternate;
}
.l-o2 {
  width: 800px; height: 800px;
  bottom: -300px; left: -200px;
  background: radial-gradient(circle, #6b46f0 0%, transparent 60%);
  animation: lOrb2 26s ease-in-out infinite alternate-reverse;
  opacity: 0.09;
}
.l-o3 {
  width: 500px; height: 500px;
  top: 40%; left: 38%;
  background: radial-gradient(circle, #f0a020 0%, transparent 60%);
  animation: lOrb3 30s ease-in-out infinite;
  opacity: 0.06;
}
.l-o4 {
  width: 350px; height: 350px;
  top: 20%; left: 10%;
  background: radial-gradient(circle, #22d88a 0%, transparent 60%);
  animation: lOrb1 23s ease-in-out infinite;
  opacity: 0.05;
}
.l-o5 {
  width: 280px; height: 280px;
  bottom: 30%; right: 15%;
  background: radial-gradient(circle, #00d2be 0%, transparent 60%);
  animation: lOrb2 32s ease-in-out infinite;
  opacity: 0.04;
}
.l-grid {
  position: fixed; inset: 0; pointer-events: none; z-index: 0;
  background-image:
    linear-gradient(rgba(0,210,190,0.03) 1px, transparent 1px),
    linear-gradient(90deg, rgba(0,210,190,0.03) 1px, transparent 1px);
  background-size: 70px 70px;
}
.l-scan {
  position: fixed; left: 0; right: 0; height: 200px;
  background: linear-gradient(180deg, transparent, rgba(0,210,190,0.03) 50%, transparent);
  pointer-events: none; z-index: 0;
  animation: lScan 8s linear infinite;
}

/* NAV */
.l-nav {
  position: fixed; top: 0; left: 0; right: 0; z-index: 1000;
  display: flex; align-items: center; justify-content: space-between;
  padding: 18px 64px;
  background: rgba(2,4,8,0.7);
  backdrop-filter: blur(32px) saturate(1.8);
  border-bottom: 1px solid rgba(0,210,190,0.08);
}
.l-brand {
  display: flex; align-items: center; gap: 12px;
  font-family: 'Syne', sans-serif;
  font-size: 1.5rem; font-weight: 800;
  color: var(--l-tx); letter-spacing: -0.5px;
}
.l-hexbadge {
  width: 36px; height: 36px; border-radius: 10px;
  background: linear-gradient(135deg, var(--l-teal), var(--l-teal2));
  display: flex; align-items: center; justify-content: center;
  font-size: 1rem; color: #020408;
  box-shadow: 0 6px 20px var(--l-glow);
  animation: lHexSpin 4s ease-in-out infinite;
}
.l-navlinks { display: flex; gap: 40px; }
.l-nl {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.6rem; letter-spacing: 2px;
  color: var(--l-tx3); cursor: default;
  transition: color 0.2s;
  text-transform: uppercase;
}
.l-nl:hover { color: var(--l-teal); }
.l-ver {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.5rem; letter-spacing: 2px;
  padding: 6px 14px; border-radius: 20px;
  background: rgba(0,210,190,0.06);
  border: 1px solid rgba(0,210,190,0.18);
  color: var(--l-teal);
}

/* HERO */
.l-hero {
  position: relative; z-index: 10;
  padding: 180px 64px 80px;
  max-width: 1400px; margin: 0 auto;
  text-align: center;
}
.l-eyebrow {
  display: inline-flex; align-items: center; gap: 10px;
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.6rem; letter-spacing: 4px;
  text-transform: uppercase; color: var(--l-teal);
  padding: 9px 22px; border-radius: 30px;
  background: rgba(0,210,190,0.05);
  border: 1px solid rgba(0,210,190,0.15);
  margin-bottom: 56px;
  animation: lFadeUp 0.8s cubic-bezier(0.2, 0.8, 0.2, 1) both;
}
.l-edot {
  width: 7px; height: 7px; border-radius: 50%;
  background: var(--l-teal);
  box-shadow: 0 0 10px var(--l-teal);
  animation: lPulse 2s ease-in-out infinite;
}
.l-headline {
  font-family: 'Syne', sans-serif;
  font-size: clamp(4rem, 12vw, 10rem);
  font-weight: 800; line-height: 0.92;
  letter-spacing: -5px; margin: 0 0 20px;
  animation: lFadeUp 0.85s 0.1s cubic-bezier(0.2, 0.8, 0.2, 1) both;
}
.l-hl1 { color: var(--l-tx); }
.l-headline em {
  font-style: italic;
  background: linear-gradient(135deg, var(--l-teal) 0%, var(--l-teal2) 50%, var(--l-amber) 100%);
  -webkit-background-clip: text; -webkit-text-fill-color: transparent;
  background-clip: text; background-size: 300%;
  animation: lFlow 5s linear infinite;
}
.l-subhead {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.75rem; letter-spacing: 6px;
  text-transform: uppercase; color: var(--l-tx3);
  margin: 0 0 48px; font-weight: 400;
  animation: lFadeUp 0.85s 0.2s cubic-bezier(0.2, 0.8, 0.2, 1) both;
}
.l-desc {
  font-size: 1.1rem; line-height: 1.8;
  color: var(--l-tx2); max-width: 640px;
  margin: 0 auto 48px; font-weight: 300;
  animation: lFadeUp 0.85s 0.3s cubic-bezier(0.2, 0.8, 0.2, 1) both;
  letter-spacing: 0.1px;
}
.l-desc strong { color: var(--l-tx); font-weight: 500; }
.l-chips {
  display: flex; flex-wrap: wrap; justify-content: center; gap: 8px;
  max-width: 900px; margin: 0 auto 52px;
  animation: lFadeUp 0.85s 0.4s cubic-bezier(0.2, 0.8, 0.2, 1) both;
}
.l-chip {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.55rem; letter-spacing: 1.5px;
  text-transform: uppercase; padding: 5px 12px;
  border-radius: 30px; font-weight: 500;
}
.l-chip-t { background: rgba(0,210,190,0.08); color: var(--l-teal); border: 1px solid rgba(0,210,190,0.2); }
.l-chip-a { background: rgba(240,160,32,0.08); color: var(--l-amber); border: 1px solid rgba(240,160,32,0.2); }
.l-chip-d { background: rgba(167,139,250,0.08); color: #a78bfa; border: 1px solid rgba(167,139,250,0.2); }

/* STAT BAR */
.l-statbar {
  display: grid; grid-template-columns: repeat(6, 1fr);
  gap: 1px; background: rgba(0,210,190,0.08);
  border-top: 1px solid rgba(0,210,190,0.1);
  border-bottom: 1px solid rgba(0,210,190,0.1);
  position: relative; z-index: 10;
  animation: lFadeUp 0.9s 0.5s cubic-bezier(0.2, 0.8, 0.2, 1) both;
}
.l-stat {
  background: rgba(4,10,18,0.8);
  padding: 30px 20px; text-align: center;
  cursor: default; transition: background 0.2s;
}
.l-stat:hover { background: rgba(0,210,190,0.04); }
.l-sv {
  font-family: 'Syne', sans-serif;
  font-size: 2.4rem; font-weight: 700;
  line-height: 1; letter-spacing: -1.5px;
}
.l-sv span { font-style: italic; }
.l-sl {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.52rem; letter-spacing: 2px;
  text-transform: uppercase; color: var(--l-tx3);
  margin-top: 6px;
}

/* TICKER */
.l-ticker-wrap {
  overflow: hidden; background: rgba(0,210,190,0.025);
  border-bottom: 1px solid rgba(0,210,190,0.07);
  padding: 12px 0; position: relative; z-index: 10;
}
.l-ticker { overflow: hidden; }
.l-ticker-inner {
  display: flex; gap: 32px; align-items: center;
  width: max-content;
  animation: lTickerScroll 40s linear infinite;
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.6rem; letter-spacing: 2px;
  text-transform: uppercase; color: var(--l-tx3);
}
.l-ticker-inner span { white-space: nowrap; }
.l-tdot { color: var(--l-teal); opacity: 0.4; }
</style>
"""

_POST_BUTTON_HTML = """
<div style="text-align:center;padding:8px 0 0;font-family:'JetBrains Mono',monospace;
  font-size:0.52rem;letter-spacing:3px;color:rgba(0,210,190,0.25);text-transform:uppercase">
  No login required · Fully local · VIT Chennai MDP 2026
</div>
<div style="height:60px"></div>
"""

_LOWER_SECTIONS = """
<style>
/* ENGINE CARDS */
.l-engines {
  padding: 80px 64px;
  max-width: 1400px; margin: 0 auto;
  position: relative; z-index: 10;
}
.l-engines-header {
  text-align: center; margin-bottom: 48px;
}
.l-eh-over {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.6rem; letter-spacing: 4px;
  text-transform: uppercase; color: rgba(0,210,190,0.5);
  margin-bottom: 12px;
}
.l-eh-title {
  font-family: 'Syne', sans-serif;
  font-size: 2.5rem; font-weight: 800;
  letter-spacing: -1.5px; color: #e4eeec;
}
.l-engine-grid {
  display: grid; grid-template-columns: repeat(3, 1fr);
  gap: 16px;
}
.l-ecard {
  background: rgba(6,16,30,0.7);
  border: 1px solid rgba(0,210,190,0.08);
  border-radius: 16px; padding: 28px;
  position: relative; overflow: hidden;
  transition: all 0.3s;
}
.l-ecard:hover {
  border-color: rgba(0,210,190,0.2);
  transform: translateY(-4px);
  box-shadow: 0 20px 60px rgba(0,0,0,0.4), 0 0 0 1px rgba(0,210,190,0.1);
}
.l-ecard::before {
  content: ''; position: absolute; top: 0; left: 0; right: 0; height: 1px;
  background: linear-gradient(90deg, transparent, var(--ac-c), transparent);
  opacity: 0; transition: opacity 0.3s;
}
.l-ecard:hover::before { opacity: 1; }
.l-etag {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.5rem; letter-spacing: 2px;
  text-transform: uppercase; padding: 4px 10px;
  border-radius: 20px; display: inline-block;
  margin-bottom: 14px;
}
.l-etag-t { background: rgba(0,210,190,0.08); color: #00d2be; border: 1px solid rgba(0,210,190,0.18); }
.l-etag-a { background: rgba(240,160,32,0.08); color: #f0a020; border: 1px solid rgba(240,160,32,0.18); }
.l-etag-v { background: rgba(167,139,250,0.08); color: #a78bfa; border: 1px solid rgba(167,139,250,0.18); }
.l-etag-g { background: rgba(34,216,138,0.08); color: #22d88a; border: 1px solid rgba(34,216,138,0.18); }
.l-ename {
  font-family: 'Syne', sans-serif;
  font-size: 1.05rem; font-weight: 700;
  color: #e4eeec; margin-bottom: 8px;
  letter-spacing: -0.3px;
}
.l-edesc {
  font-size: 0.78rem; color: rgba(160,200,190,0.55);
  line-height: 1.7; font-family: 'Inter', sans-serif;
}
.l-escore {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.55rem; color: rgba(0,210,190,0.35);
  margin-top: 12px; letter-spacing: 1px;
}

/* QUOTE SECTION */
.l-quote {
  padding: 80px 64px;
  border-top: 1px solid rgba(0,210,190,0.06);
  border-bottom: 1px solid rgba(0,210,190,0.06);
  text-align: center;
  background: linear-gradient(135deg, rgba(0,210,190,0.015) 0%, transparent 50%, rgba(240,160,32,0.015) 100%);
  position: relative; z-index: 10;
}
.l-ql {
  font-family: 'Syne', sans-serif;
  font-size: clamp(1.4rem, 3.5vw, 2.4rem);
  font-style: italic; font-weight: 600;
  color: rgba(228,238,236,0.7);
  max-width: 900px; margin: 0 auto 20px;
  line-height: 1.4; letter-spacing: -0.5px;
}
.l-qa {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.58rem; letter-spacing: 3px;
  text-transform: uppercase;
  color: rgba(0,210,190,0.4);
}

/* WORKFLOW */
.l-workflow {
  padding: 80px 64px;
  max-width: 1400px; margin: 0 auto;
  position: relative; z-index: 10;
}
.l-wf-header {
  text-align: center; margin-bottom: 56px;
}
.l-wf-steps {
  display: grid; grid-template-columns: repeat(4, 1fr);
  gap: 2px; position: relative;
}
.l-wf-steps::before {
  content: ''; position: absolute;
  top: 28px; left: 5%; right: 5%; height: 1px;
  background: linear-gradient(90deg, transparent, rgba(0,210,190,0.3), transparent);
}
.l-wf-step {
  background: rgba(6,16,30,0.6);
  border: 1px solid rgba(0,210,190,0.07);
  border-radius: 14px; padding: 28px 24px;
  text-align: center; position: relative;
  transition: all 0.3s;
}
.l-wf-step:hover {
  border-color: rgba(0,210,190,0.18);
  background: rgba(0,210,190,0.03);
  transform: translateY(-4px);
}
.l-wf-num {
  width: 44px; height: 44px; border-radius: 50%;
  background: rgba(0,210,190,0.06);
  border: 1.5px solid rgba(0,210,190,0.18);
  display: flex; align-items: center; justify-content: center;
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.8rem; font-weight: 600; color: #00d2be;
  margin: 0 auto 16px;
}
.l-wf-title {
  font-family: 'Syne', sans-serif;
  font-size: 0.9rem; font-weight: 700;
  color: #e4eeec; margin-bottom: 8px;
}
.l-wf-desc {
  font-size: 0.72rem; color: rgba(160,200,190,0.5);
  line-height: 1.6; font-family: 'Inter', sans-serif;
}

/* TECH STACK */
.l-stack {
  padding: 60px 64px 80px;
  max-width: 1400px; margin: 0 auto;
  position: relative; z-index: 10;
}
.l-stack-row {
  display: flex; gap: 10px; flex-wrap: wrap;
  justify-content: center; margin-top: 32px;
}
.l-stack-item {
  background: rgba(6,16,30,0.8);
  border: 1px solid rgba(0,210,190,0.08);
  border-radius: 10px;
  padding: 10px 18px;
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.65rem; letter-spacing: 1px;
  color: rgba(160,200,190,0.5);
  transition: all 0.2s;
}
.l-stack-item:hover {
  border-color: rgba(0,210,190,0.2);
  color: rgba(0,210,190,0.8);
}

/* CTA SECTION */
.l-cta {
  padding: 60px 64px 100px;
  text-align: center;
  position: relative; z-index: 10;
}
.l-cta-title {
  font-family: 'Syne', sans-serif;
  font-size: 1.2rem; font-weight: 700;
  color: rgba(228,238,236,0.6);
  margin-bottom: 8px; letter-spacing: -0.3px;
}
.l-cta-sub {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.58rem; letter-spacing: 2px;
  text-transform: uppercase;
  color: rgba(0,210,190,0.3);
  margin-bottom: 32px;
}

/* FOOTER */
.l-foot {
  border-top: 1px solid rgba(0,210,190,0.06);
  padding: 28px 64px;
  display: flex; align-items: center; justify-content: space-between;
  position: relative; z-index: 10;
}
.l-foot-brand {
  font-family: 'Syne', sans-serif;
  font-size: 1.1rem; font-weight: 800;
  color: rgba(228,238,236,0.4);
  letter-spacing: -0.5px;
}
.l-foot-refs {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.5rem; letter-spacing: 1.5px;
  color: rgba(0,210,190,0.2);
  text-align: right; line-height: 2;
}
</style>

<!-- ENGINE CARDS -->
<section class="l-engines">
  <div class="l-engines-header">
    <div class="l-eh-over">Intelligence Architecture</div>
    <div class="l-eh-title">9 Tiered Analysis Engines</div>
  </div>
  <div class="l-engine-grid">

    <div class="l-ecard" style="--ac-c: #00d2be">
      <div class="l-etag l-etag-t">Core Engine</div>
      <div class="l-ename">Vanguard Core</div>
      <div class="l-edesc">21-parameter ADMET screening covering Lipinski, BOILED-EGG, QED, SA Score, CNS MPO, ESOL solubility, and full CYP/hERG/PAINS safety panel.</div>
      <div class="l-escore">PARAMETERS: 21+ · RDKIT NATIVE</div>
    </div>

    <div class="l-ecard" style="--ac-c: #f0a020">
      <div class="l-etag l-etag-a">Tier 3</div>
      <div class="l-ename">Hyper-Zenith v50</div>
      <div class="l-edesc">Advanced ADME predictors: Caco-2 permeability, P-gp substrate alerts, skin permeability (LogKp), DILI risk, phospholipidosis index, and metabolic half-life.</div>
      <div class="l-escore">PARAMETERS: 500+ · PRECISION TIER</div>
    </div>

    <div class="l-ecard" style="--ac-c: #a78bfa">
      <div class="l-etag l-etag-v">Tier 5</div>
      <div class="l-ename">Omni-Science v20</div>
      <div class="l-edesc">50+ new molecular descriptors spanning topology, physicochemistry, elemental ratios, SAR intelligence, and IP originality scoring with CNS MPO v2.</div>
      <div class="l-escore">PARAMETERS: 2000+ · MEGA ENGINE</div>
    </div>

    <div class="l-ecard" style="--ac-c: #22d88a">
      <div class="l-etag l-etag-g">Tier 7</div>
      <div class="l-ename">Celestial v1000</div>
      <div class="l-edesc">Mechanistic PBPK kinetics with Ka, CLint, Kp tissue partitioning. Quantum electronic descriptors (QUED), Saagar hazard registry, and SHAP explainability.</div>
      <div class="l-escore">PARAMETERS: 10,000+ · SUPREME TIER</div>
    </div>

    <div class="l-ecard" style="--ac-c: #f0a020">
      <div class="l-etag l-etag-a">Tier 8</div>
      <div class="l-ename">Xenon-God v5000</div>
      <div class="l-edesc">50,000+ parameter multiverse analysis: quantum orbital dynamics, hydration energies, epigenetic hazard scan, retrosynthetic difficulty index (RDI), and BBB flux tags.</div>
      <div class="l-escore">PARAMETERS: 50,000+ · MULTIVERSE</div>
    </div>

    <div class="l-ecard" style="--ac-c: #00d2be">
      <div class="l-etag l-etag-t">Tier 9</div>
      <div class="l-ename">Aether-Primality v10000</div>
      <div class="l-edesc">God-mode engine with 100,000+ feature tensors. Tissue-specific PBPK permeability, nanotoxicity screening, carbon footprint analysis, and quantum Aether motifs.</div>
      <div class="l-escore">PARAMETERS: 100,000+ · GOD ENGINE</div>
    </div>

  </div>
</section>

<!-- QUOTE -->
<div class="l-quote">
  <div class="l-ql">"Targeted certainty in a multiverse of chemical possibilities."</div>
  <div class="l-qa">Omega Protocol Engaged · System Omnipotent · VIT Chennai 2026</div>
</div>

<!-- WORKFLOW -->
<section class="l-workflow">
  <div class="l-wf-header">
    <div class="l-eh-over">How it works</div>
    <div class="l-eh-title">Four-Stage Discovery Pipeline</div>
  </div>
  <div class="l-wf-steps">
    <div class="l-wf-step">
      <div class="l-wf-num">01</div>
      <div class="l-wf-title">Input SMILES</div>
      <div class="l-wf-desc">Paste comma-separated SMILES strings, upload CSV/Excel, or load directly from SDF/MOL files.</div>
    </div>
    <div class="l-wf-step">
      <div class="l-wf-num">02</div>
      <div class="l-wf-title">Deep ADMET Scan</div>
      <div class="l-wf-desc">All 9 engine tiers run in parallel. 100,000+ parameters evaluated per compound in seconds.</div>
    </div>
    <div class="l-wf-step">
      <div class="l-wf-num">03</div>
      <div class="l-wf-title">Lead Scoring</div>
      <div class="l-wf-desc">ChemoScore v1.0 synthesizes all parameters into a single 0–100 lead quality index with full grade assignment.</div>
    </div>
    <div class="l-wf-step">
      <div class="l-wf-num">04</div>
      <div class="l-wf-title">Export & Report</div>
      <div class="l-wf-desc">Download professional reports in CSV, HTML, TXT, or PDF-ready format. 202 references included.</div>
    </div>
  </div>
</section>

<!-- TECH STACK -->
<section class="l-stack">
  <div class="l-wf-header">
    <div class="l-eh-over">Built With</div>
    <div class="l-eh-title">Production-Grade Stack</div>
  </div>
  <div class="l-stack-row">
    <div class="l-stack-item">RDKit</div>
    <div class="l-stack-item">Streamlit</div>
    <div class="l-stack-item">Plotly</div>
    <div class="l-stack-item">Anthropic Claude</div>
    <div class="l-stack-item">NumPy</div>
    <div class="l-stack-item">Pandas</div>
    <div class="l-stack-item">Python 3.11+</div>
    <div class="l-stack-item">PubChem REST API</div>
    <div class="l-stack-item">Morgan Fingerprints ECFP4</div>
    <div class="l-stack-item">MMFF94 Forcefield</div>
    <div class="l-stack-item">PAINS Catalog</div>
    <div class="l-stack-item">Brenk Filter</div>
    <div class="l-stack-item">Murcko Scaffolds</div>
    <div class="l-stack-item">FilterCatalog</div>
    <div class="l-stack-item">SA Score (Ertl 2009)</div>
    <div class="l-stack-item">BOILED-EGG (Daina 2016)</div>
  </div>
</section>

<!-- CTA SECTION -->
<div class="l-cta">
  <div class="l-cta-title">Ready to Screen Your Compounds?</div>
  <div class="l-cta-sub">21+ ADMET parameters · 9 engine tiers · AI-powered insights</div>
</div>

<!-- FOOTER -->
<div class="l-foot">
  <div class="l-foot-brand">ChemoFilter</div>
  <div class="l-foot-refs">
    BOILED-EGG [DAINA 2016] · LIPINSKI [2001] · ESOL [DELANEY 2004]<br>
    QED [BICKERTON 2012] · SA SCORE [ERTL 2009] · CNS MPO [WAGER 2010]<br>
    PAINS [BAELL 2010] · RDKIT [LANDRUM] · 202 TOTAL REFERENCES
  </div>
</div>
"""
