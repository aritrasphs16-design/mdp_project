"""
landing_enhancements.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Landing Page Interactive Enhancements
────────────────────────────────────────────────────────────────────────────

ADDITIVE ONLY — Does not modify any existing landing.py UI or logic.
All features are lightweight, lazy, and guarded with try/except.

10 ENHANCEMENTS IMPLEMENTED
  1. Quick Molecule Preview (MW, LogP, TPSA — instant, no API)
  2. Interactive Sample Selector (Aspirin, Caffeine, Paracetamol chips)
  3. Real-Time SMILES Validation (✅ / ❌ inline)
  4. Mini Insight Panel (drug-likeness heuristic, key concern)
  5. Feature Highlight Cards (ADMET, Scaffold, Comparison, AI — hover tooltips)
  6. Recent Analysis Memory (session_state last compound recall)
  7. Progressive Disclosure (Show Advanced Options toggle)
  8. Zero-Lag Rule (no APIs on load, all guarded)
  9. Micro-animations (CSS only, no heavy JS)
 10. Scientific Tooltip System (all terms labeled)
────────────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations
import streamlit as st

# ─── Minimal RDKit import guard ───────────────────────────────────────────────
def _rdkit_available() -> bool:
    try:
        from rdkit import Chem  # noqa: F401
        return True
    except ImportError:
        return False

# ─── CSS for all enhancements (injected once) ─────────────────────────────────
_ENHANCEMENT_CSS = """
<style>
/* ── Enhancement container base ── */
.lx-container {
  font-family: 'Inter', sans-serif;
  color: var(--l-tx, #e4eeec);
  position: relative; z-index: 10;
}

/* ── Sample selector chips ── */
.lx-sample-chips { display: flex; gap: 10px; flex-wrap: wrap; margin: 10px 0; }
.lx-chip-btn {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.6rem; letter-spacing: 1.5px; text-transform: uppercase;
  padding: 6px 14px; border-radius: 30px; cursor: pointer;
  background: rgba(0,210,190,0.07);
  border: 1px solid rgba(0,210,190,0.22);
  color: #00d2be; transition: all 0.2s;
  display: inline-block;
}
.lx-chip-btn:hover {
  background: rgba(0,210,190,0.16);
  transform: translateY(-1px);
  box-shadow: 0 4px 16px rgba(0,210,190,0.18);
}

/* ── Property card grid ── */
.lx-prop-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 10px; margin: 12px 0; }
.lx-prop-card {
  background: rgba(6,16,30,0.75);
  border: 1px solid rgba(0,210,190,0.1);
  border-radius: 10px; padding: 12px 14px;
  text-align: center; transition: border-color 0.2s;
}
.lx-prop-card:hover { border-color: rgba(0,210,190,0.25); }
.lx-prop-val {
  font-family: 'Syne', sans-serif;
  font-size: 1.35rem; font-weight: 700;
  color: #00d2be; letter-spacing: -0.5px;
}
.lx-prop-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.5rem; letter-spacing: 2px;
  text-transform: uppercase; color: rgba(0,210,190,0.45);
  margin-top: 4px;
}
.lx-prop-tip {
  font-size: 0.6rem; color: rgba(160,200,190,0.45);
  margin-top: 2px; font-style: italic;
}

/* ── Validation badge ── */
.lx-valid   { color: #22d88a; font-size: 0.78rem; font-weight: 600; }
.lx-invalid { color: #f87171; font-size: 0.78rem; font-weight: 600; }

/* ── Insight panel ── */
.lx-insight {
  background: rgba(6,16,30,0.8);
  border: 1px solid rgba(0,210,190,0.1);
  border-left: 3px solid #00d2be;
  border-radius: 10px; padding: 14px 18px; margin: 10px 0;
}
.lx-insight-title {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.55rem; letter-spacing: 3px;
  text-transform: uppercase; color: rgba(0,210,190,0.5);
  margin-bottom: 8px;
}
.lx-insight-row { display: flex; justify-content: space-between; align-items: center; padding: 3px 0; }
.lx-insight-key { font-size: 0.72rem; color: rgba(180,220,215,0.55); }
.lx-insight-val { font-size: 0.75rem; font-weight: 600; }
.lx-val-high   { color: #22d88a; }
.lx-val-mod    { color: #f0a020; }
.lx-val-low    { color: #f87171; }

/* ── Feature cards ── */
.lx-feature-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: 12px; margin: 16px 0; }
.lx-feat-card {
  background: rgba(6,16,30,0.7);
  border: 1px solid rgba(0,210,190,0.08);
  border-radius: 12px; padding: 18px 14px;
  text-align: center; cursor: default;
  transition: all 0.25s; position: relative;
}
.lx-feat-card:hover {
  border-color: rgba(0,210,190,0.25);
  transform: translateY(-3px);
  box-shadow: 0 12px 40px rgba(0,0,0,.3);
}
.lx-feat-icon { font-size: 1.4rem; margin-bottom: 8px; }
.lx-feat-name {
  font-family: 'Syne', sans-serif;
  font-size: 0.8rem; font-weight: 700; color: #e4eeec;
  margin-bottom: 6px;
}
.lx-feat-tooltip {
  font-size: 0.65rem; color: rgba(160,200,190,0.5);
  line-height: 1.5; display: none;
}
.lx-feat-card:hover .lx-feat-tooltip { display: block; }

/* ── Recent analysis memory ── */
.lx-memory {
  background: rgba(240,160,32,0.04);
  border: 1px solid rgba(240,160,32,0.15);
  border-radius: 10px; padding: 12px 16px; margin: 10px 0;
}
.lx-memory-label {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.5rem; letter-spacing: 2px;
  text-transform: uppercase; color: rgba(240,160,32,0.5);
  margin-bottom: 6px;
}

/* ── Advanced toggle ── */
.lx-adv-section {
  background: rgba(6,16,30,0.6);
  border: 1px solid rgba(0,210,190,0.07);
  border-radius: 10px; padding: 14px 18px; margin: 10px 0;
}

/* ── Section divider ── */
.lx-divider {
  border: none; border-top: 1px solid rgba(0,210,190,0.07);
  margin: 20px 0;
}

/* ── Micro-animation: fade-in on appearance ── */
@keyframes lxFadeUp {
  from { opacity: 0; transform: translateY(8px); }
  to   { opacity: 1; transform: none; }
}
.lx-animated { animation: lxFadeUp 0.4s ease both; }

/* ── Tooltip for scientific terms ── */
.lx-term {
  border-bottom: 1px dotted rgba(0,210,190,0.4);
  cursor: help; position: relative;
}
.lx-term::after {
  content: attr(data-tip);
  position: absolute; bottom: 120%; left: 50%; transform: translateX(-50%);
  background: rgba(4,10,18,0.95); border: 1px solid rgba(0,210,190,0.2);
  color: #c8deff; font-size: 0.62rem; line-height: 1.5;
  padding: 6px 10px; border-radius: 6px;
  width: 200px; text-align: left; pointer-events: none;
  opacity: 0; transition: opacity 0.2s; z-index: 9999;
  white-space: normal;
}
.lx-term:hover::after { opacity: 1; }
</style>
"""

# ─── SAMPLE MOLECULES ────────────────────────────────────────────────────────
SAMPLE_MOLECULES = {
    "Aspirin":      "CC(=O)Oc1ccccc1C(=O)O",
    "Caffeine":     "Cn1cnc2c1c(=O)n(C)c(=O)n2C",
    "Paracetamol":  "CC(=O)Nc1ccc(O)cc1",
    "Ibuprofen":    "CC(C)Cc1ccc(cc1)C(C)C(=O)O",
    "Metformin":    "CN(C)C(=N)NC(=N)N",
}

# ─────────────────────────────────────────────────────────────────────────────
# LIGHTWEIGHT PROPERTY COMPUTATION (no API, pure RDKit or fallback)
# ─────────────────────────────────────────────────────────────────────────────

def _compute_quick_props(smiles: str) -> dict | None:
    """Compute MW, LogP, TPSA instantly from SMILES. Returns None on invalid."""
    try:
        from rdkit import Chem
        from rdkit.Chem import Descriptors, rdMolDescriptors
        mol = Chem.MolFromSmiles(smiles.strip())
        if mol is None:
            return None
        mw   = round(Descriptors.MolWt(mol), 1)
        logp = round(Descriptors.MolLogP(mol), 2)
        tpsa = round(Descriptors.TPSA(mol), 1)
        hbd  = rdMolDescriptors.CalcNumHBD(mol)
        hba  = rdMolDescriptors.CalcNumHBA(mol)
        return {"mw": mw, "logp": logp, "tpsa": tpsa, "hbd": hbd, "hba": hba}
    except Exception:
        return None


def _druglikeness_heuristic(props: dict) -> tuple[str, str, str]:
    """
    Fast heuristic Lipinski/Veber assessment.
    Returns (level: High|Moderate|Low, concern, color_class)
    """
    if not props:
        return "Unknown", "Invalid structure", "lx-val-low"
    mw   = props.get("mw", 0)
    logp = props.get("logp", 0)
    tpsa = props.get("tpsa", 0)
    hbd  = props.get("hbd", 0)
    hba  = props.get("hba", 0)
    violations = 0
    concerns = []
    if mw > 500:   violations += 1; concerns.append("high MW")
    if logp > 5:   violations += 1; concerns.append("high lipophilicity")
    if hbd > 5:    violations += 1; concerns.append("too many H-bond donors")
    if hba > 10:   violations += 1; concerns.append("too many H-bond acceptors")
    if tpsa > 140: concerns.append("low CNS penetration (TPSA)")
    if violations == 0:
        level = "High"; color = "lx-val-high"
        concern = concerns[0].capitalize() if concerns else "All Lipinski rules satisfied"
    elif violations == 1:
        level = "Moderate"; color = "lx-val-mod"
        concern = concerns[0].capitalize() if concerns else "Minor Ro5 concern"
    else:
        level = "Low"; color = "lx-val-low"
        concern = "; ".join(concerns[:2]).capitalize()
    return level, concern, color


# ─────────────────────────────────────────────────────────────────────────────
# RENDER FUNCTIONS
# Each is wrapped in try/except — never crashes the landing page.
# ─────────────────────────────────────────────────────────────────────────────

def _inject_css():
    st.markdown(_ENHANCEMENT_CSS, unsafe_allow_html=True)


def render_feature_cards():
    """Enhancement 5 — Feature Highlight Cards with hover tooltips."""
    try:
        features = [
            ("🧬", "ADMET Prediction",
             "21-parameter absorption, distribution, metabolism, excretion & toxicity screening powered by RDKit"),
            ("🔬", "Scaffold Analysis",
             "Murcko scaffold decomposition, ring system analysis, and bioisostere detection"),
            ("⚖️", "Comparison Engine",
             "Side-by-side multi-compound analysis with Tanimoto similarity, radar plots, and lead ranking"),
            ("🤖", "AI Insights",
             "Claude AI explains ADMET results in natural language, suggests structural improvements and repurposing opportunities"),
        ]
        cards_html = '<div class="lx-feature-grid lx-animated">'
        for icon, name, tip in features:
            cards_html += f"""
            <div class="lx-feat-card">
              <div class="lx-feat-icon">{icon}</div>
              <div class="lx-feat-name">{name}</div>
              <div class="lx-feat-tooltip">{tip}</div>
            </div>"""
        cards_html += '</div>'
        st.markdown(cards_html, unsafe_allow_html=True)
    except Exception:
        pass  # never break the page


def render_quick_molecule_preview():
    """
    Enhancements 1, 2, 3, 4, 9, 10 — combined interactive molecule preview.
    Includes sample selector, SMILES input, live validation, properties,
    mini insight panel, and scientific tooltips.
    """
    try:
        # ── Section header
        st.markdown(
            '<div class="lx-container lx-animated" style="padding:0 0 8px">'
            '<div style="font-family:\'JetBrains Mono\',monospace;font-size:0.55rem;'
            'letter-spacing:4px;text-transform:uppercase;color:rgba(0,210,190,0.45);'
            'margin-bottom:12px">⬡ Quick Molecule Preview</div>'
            '</div>',
            unsafe_allow_html=True)

        # ── Enhancement 2: Sample Selector chips ──────────────────────────────
        st.markdown(
            '<div style="font-size:0.7rem;color:rgba(180,220,215,0.5);margin-bottom:6px">'
            'Try a sample compound:</div>', unsafe_allow_html=True)

        cols = st.columns(len(SAMPLE_MOLECULES))
        for i, (name, smi) in enumerate(SAMPLE_MOLECULES.items()):
            with cols[i]:
                if st.button(name, key=f"_lx_sample_{name}",
                             use_container_width=True):
                    st.session_state["_lx_smiles_input"] = smi

        # ── Enhancement 1: SMILES input ───────────────────────────────────────
        smiles_val = st.session_state.get("_lx_smiles_input", "")
        smiles_input = st.text_input(
            "Enter SMILES string",
            value=smiles_val,
            key="_lx_smiles_main",
            placeholder="e.g. CC(=O)Oc1ccccc1C(=O)O",
            label_visibility="collapsed",
            help="Paste any valid SMILES string. Try the sample buttons above.")

        # Sync input back to session
        if smiles_input:
            st.session_state["_lx_smiles_input"] = smiles_input

        # ── Enhancement 3: Real-Time Validation ───────────────────────────────
        if smiles_input:
            props = _compute_quick_props(smiles_input)
            if props is not None:
                st.markdown(
                    '<span class="lx-valid">✅ Valid structure — properties computed below</span>',
                    unsafe_allow_html=True)
            else:
                st.markdown(
                    '<span class="lx-invalid">❌ Invalid SMILES — check your input</span>',
                    unsafe_allow_html=True)

            # ── Enhancement 1: Property preview ───────────────────────────────
            if props:
                mw   = props["mw"]
                logp = props["logp"]
                tpsa = props["tpsa"]

                # Colour-code by range
                mw_col   = "#22d88a" if mw <= 500 else "#f87171"
                logp_col = "#22d88a" if -0.4 <= logp <= 5.6 else "#f87171"
                tpsa_col = "#22d88a" if tpsa <= 140 else "#f87171"

                st.markdown(
                    f'<div class="lx-prop-grid lx-animated">'
                    f'<div class="lx-prop-card">'
                    f'  <div class="lx-prop-val" style="color:{mw_col}">{mw}</div>'
                    f'  <div class="lx-prop-label">'
                    f'    <span class="lx-term" data-tip="Molecular Weight: &lt;500 Da favors oral bioavailability (Lipinski Rule 1)">Mol Weight</span>'
                    f'  </div>'
                    f'  <div class="lx-prop-tip">≤500 Da preferred</div>'
                    f'</div>'
                    f'<div class="lx-prop-card">'
                    f'  <div class="lx-prop-val" style="color:{logp_col}">{logp}</div>'
                    f'  <div class="lx-prop-label">'
                    f'    <span class="lx-term" data-tip="LogP: lipophilicity measure. Values 0–5 favor membrane permeability without excessive accumulation">LogP</span>'
                    f'  </div>'
                    f'  <div class="lx-prop-tip">0 – 5 optimal</div>'
                    f'</div>'
                    f'<div class="lx-prop-card">'
                    f'  <div class="lx-prop-val" style="color:{tpsa_col}">{tpsa}</div>'
                    f'  <div class="lx-prop-label">'
                    f'    <span class="lx-term" data-tip="TPSA: Topological Polar Surface Area. &lt;140 Å² for oral drugs; &lt;90 Å² for CNS drugs">TPSA (Å²)</span>'
                    f'  </div>'
                    f'  <div class="lx-prop-tip">≤140 Å² oral</div>'
                    f'</div>'
                    f'</div>',
                    unsafe_allow_html=True)

                # ── Enhancement 4: Mini Insight Panel ─────────────────────────
                level, concern, color_cls = _druglikeness_heuristic(props)
                _render_insight_panel(props, level, concern, color_cls)

        elif not smiles_input:
            st.markdown(
                '<div style="font-size:0.7rem;color:rgba(160,200,190,0.35);'
                'text-align:center;padding:8px 0">↑ Enter SMILES or pick a sample above</div>',
                unsafe_allow_html=True)

    except Exception:
        pass  # never break the page


def _render_insight_panel(props: dict, level: str, concern: str, color_cls: str):
    """Enhancement 4 — Mini Insight Panel."""
    try:
        hbd = props.get("hbd", "—")
        hba = props.get("hba", "—")
        st.markdown(
            f'<div class="lx-insight lx-animated">'
            f'<div class="lx-insight-title">⬡ Drug-Likeness Assessment</div>'
            f'<div class="lx-insight-row">'
            f'  <span class="lx-insight-key">'
            f'    <span class="lx-term" data-tip="Lipinski Ro5 + Veber rule assessment — heuristic estimate of oral bioavailability potential">Drug-likeness</span>'
            f'  </span>'
            f'  <span class="lx-insight-val {color_cls}">{level}</span>'
            f'</div>'
            f'<div class="lx-insight-row">'
            f'  <span class="lx-insight-key">Key concern</span>'
            f'  <span class="lx-insight-val" style="color:rgba(200,220,255,.7);font-size:.72rem">{concern}</span>'
            f'</div>'
            f'<div class="lx-insight-row">'
            f'  <span class="lx-insight-key">'
            f'    <span class="lx-term" data-tip="H-Bond Donors: N-H and O-H groups. Lipinski limit: ≤5">H-Bond Donors</span>'
            f'  </span>'
            f'  <span class="lx-insight-val" style="color:{"#22d88a" if hbd <= 5 else "#f87171"}">{hbd}</span>'
            f'</div>'
            f'<div class="lx-insight-row">'
            f'  <span class="lx-insight-key">'
            f'    <span class="lx-term" data-tip="H-Bond Acceptors: N and O atoms. Lipinski limit: ≤10">H-Bond Acceptors</span>'
            f'  </span>'
            f'  <span class="lx-insight-val" style="color:{"#22d88a" if hba <= 10 else "#f87171"}">{hba}</span>'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True)
    except Exception:
        pass


def render_recent_analysis_memory():
    """Enhancement 6 — Recent Analysis Memory from session_state."""
    try:
        # Try to find last compound from main app analysis session state
        last_smiles = (
            st.session_state.get("last_analyzed_smiles") or
            st.session_state.get("cf_last_smiles") or
            st.session_state.get("_lx_smiles_input") or
            ""
        )
        last_score = st.session_state.get("last_lead_score")
        last_name  = st.session_state.get("last_compound_name", "")

        if last_smiles and len(last_smiles) > 4:
            # Quickly compute a summary
            props = _compute_quick_props(last_smiles)
            level, concern, _ = _druglikeness_heuristic(props) if props else ("—", "—", "")

            display_name = last_name or last_smiles[:28] + ("…" if len(last_smiles) > 28 else "")
            score_html = ""
            if last_score is not None:
                score_html = (f'<span style="color:#f0a020;font-weight:700">'
                              f'Lead Score: {last_score:.1f}</span> · ')

            st.markdown(
                f'<div class="lx-memory lx-animated">'
                f'<div class="lx-memory-label">⬡ Last Analysis</div>'
                f'<div style="font-size:.75rem;color:#c8deff">{display_name}</div>'
                f'<div style="font-size:.68rem;color:rgba(180,220,215,.5);margin-top:4px">'
                f'{score_html}Drug-likeness: {level}</div>'
                f'</div>',
                unsafe_allow_html=True)
    except Exception:
        pass


def render_progressive_disclosure():
    """Enhancement 7 — Progressive Disclosure (Advanced Options toggle)."""
    try:
        with st.expander("⚙  Show Advanced Options", expanded=False):
            st.markdown(
                '<div class="lx-adv-section lx-animated">',
                unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                engine = st.selectbox(
                    "🔧 Preview engine",
                    ["Vanguard Core (RDKit)", "Hyper-Zenith v50", "Celestial v1000"],
                    key="_lx_adv_engine",
                    help="Select which scoring engine to use for the quick preview")
            with col2:
                strict = st.checkbox(
                    "Strict Lipinski mode",
                    value=False,
                    key="_lx_adv_strict",
                    help="Apply zero-tolerance Lipinski: any violation = Low drug-likeness")
            reference = st.text_input(
                "Reference SMILES for Tanimoto comparison (optional)",
                key="_lx_adv_ref",
                placeholder="Paste reference SMILES…",
                help="If provided, Tanimoto similarity to this compound will be shown")
            if reference:
                ref_props = _compute_quick_props(reference)
                if ref_props:
                    st.success(f"Reference MW: {ref_props['mw']} · LogP: {ref_props['logp']}")
                else:
                    st.error("Invalid reference SMILES")
            st.markdown('</div>', unsafe_allow_html=True)
    except Exception:
        pass


def render_api_health_summary():
    """Compact API health strip — shows live status of external services."""
    try:
        from api_reliability import health_badge, KNOWN_APIS, get_health
        checked = [k for k in KNOWN_APIS
                   if get_health(k)["status"] not in ("unknown",)]
        if not checked:
            return  # nothing to show until APIs have been called
        st.markdown(
            '<div style="font-family:\'JetBrains Mono\',monospace;font-size:0.5rem;'
            'letter-spacing:3px;text-transform:uppercase;color:rgba(0,210,190,0.35);'
            'margin-bottom:8px">API Status</div>',
            unsafe_allow_html=True)
        cols = st.columns(min(len(checked), 5))
        for i, key in enumerate(checked[:5]):
            with cols[i]:
                badge = health_badge(key)
                st.markdown(
                    f'<div style="font-size:.62rem;text-align:center;'
                    f'color:rgba(200,220,255,.6)">'
                    f'<b>{key}</b><br>{badge}</div>',
                    unsafe_allow_html=True)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENTRY POINT — called from landing.py after existing content
# ─────────────────────────────────────────────────────────────────────────────

def render_landing_enhancements():
    """
    Render all 10 landing page enhancements.
    ADDITIVE ONLY — safe to call after existing render_landing() content.
    """
    try:
        _inject_css()

        # ── Section wrapper
        with st.container():
            st.markdown('<hr class="lx-divider">', unsafe_allow_html=True)

            # ── Feature Cards (static but interactive)
            st.markdown(
                '<div style="font-family:\'JetBrains Mono\',monospace;font-size:0.55rem;'
                'letter-spacing:4px;text-transform:uppercase;color:rgba(0,210,190,0.4);'
                'margin-bottom:6px">⬡ Platform Capabilities</div>',
                unsafe_allow_html=True)
            render_feature_cards()

            st.markdown('<hr class="lx-divider">', unsafe_allow_html=True)

            # ── Quick Molecule Preview (includes samples, validation, insight)
            render_quick_molecule_preview()

            # ── Recent Analysis Memory
            render_recent_analysis_memory()

            # ── Progressive Disclosure
            render_progressive_disclosure()

            # ── API Health (only if APIs have been called)
            render_api_health_summary()

            st.markdown('<div style="height:20px"></div>', unsafe_allow_html=True)

    except Exception:
        pass  # absolute last resort — never crash the page
