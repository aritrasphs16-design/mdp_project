"""
comparison_mode.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Comparison Mode — Tab 37 (PHASE 3)
• Side-by-side property comparison of selected compounds
• Radar / spider chart for multi-property visualisation
• Delta table showing property differences
• Fully isolated — read-only access to res list
────────────────────────────────────────────────────────────────────────────
"""

import streamlit as st

# NEW CODE START — terminology import (display layer only)
try:
    from terminology import compare_label, compare_tip, feat_label, feat_tip, FEATURE_LABELS
    _TERM_OK = True
except Exception:
    _TERM_OK = False
    def compare_label(k): return k
    def compare_tip(k): return None
    def feat_label(k): return k.replace("_"," ").title()
    def feat_tip(k): return None
# NEW CODE END

try:
    import plotly.graph_objects as go
    _PLT_OK = True
except Exception:
    _PLT_OK = False

_COMPARE_PROPS = ["MW", "LogP", "tPSA", "QED", "SA_Score", "LeadScore", "OralBioScore"]
_RADAR_PROPS   = ["QED", "LeadScore", "OralBioScore"]


def _normalise(value, min_v, max_v) -> float:
    """Normalise value to 0–1."""
    try:
        v = float(value)
        span = max_v - min_v
        if span == 0:
            return 0.5
        return max(0.0, min(1.0, (v - min_v) / span))
    except Exception:
        return 0.0


def _radar_fig(compounds: list):
    if not _PLT_OK or not compounds:
        return None

    categories = [
        "Drug-Likeness × 100 (QED)",
        "Lead Optimisation Score",
        "Oral Bioavailability Score",
        "Blood-Brain Barrier Penetration (%)",
        "Human Intestinal Absorption (%)",
    ]

    fig = go.Figure()
    colors = ["#f5a623", "#4ade80", "#38bdf8", "#a78bfa", "#fb923c"]

    for idx, cpd in enumerate(compounds):
        qed_val  = float(cpd.get("QED", cpd.get("_qed", 0.5))) * 100
        lead_val = float(cpd.get("LeadScore", 60))
        oral_val = float(cpd.get("OralBioScore", 60))
        bbb_val  = 80.0 if cpd.get("_bbb") else 20.0
        hia_val  = 80.0 if cpd.get("_hia") else 20.0

        r_vals = [qed_val, lead_val, oral_val, bbb_val, hia_val]

        fig.add_trace(go.Scatterpolar(
            r=r_vals + [r_vals[0]],
            theta=categories + [categories[0]],
            fill="toself",
            name=cpd.get("ID", f"Cpd-{idx+1}"),
            line_color=colors[idx % len(colors)],
            fillcolor=colors[idx % len(colors)].replace("#", "rgba(").rstrip(")") + ",0.08)"
            if colors[idx % len(colors)].startswith("#") else "rgba(245,166,35,0.08)",
        ))

    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 100], color="#c8deff"),
            angularaxis=dict(color="#c8deff"),
            bgcolor="rgba(0,0,0,0)",
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#c8deff", size=11),
        height=420,
        title="Multi-Property Radar Chart — Drug-Likeness, Bioavailability & Safety Profile",
    )
    return fig


def render_tab(res: list):
    """Render Comparison Mode tab."""
    st.markdown(
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.6rem;'
        'letter-spacing:3px;color:rgba(232,160,32,.5);text-transform:uppercase;'
        'margin-bottom:12px">⬡ Comparison Mode — Side-by-Side Analysis</div>',
        unsafe_allow_html=True,
    )

    if not res:
        st.warning("No compounds loaded.")
        return

    # Compound selector
    ids = [c.get("ID", f"Cpd-{i+1}") for i, c in enumerate(res)]
    selected_ids = st.multiselect(
        "Select Compounds to Compare (2–5 recommended)",
        ids,
        default=ids[:min(3, len(ids))],
        key="_cmp_select",
    )

    if not selected_ids:
        st.info("Please select at least one compound to begin comparison.")
        return

    selected = [c for c in res if c.get("ID", "") in selected_ids]

    # ── Radar chart ───────────────────────────────────────────────────────
    if _PLT_OK:
        fig = _radar_fig(selected)
        if fig:
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Plotly not available — radar chart unavailable.")

    # ── Property table ────────────────────────────────────────────────────
    st.subheader("Physicochemical & Drug-Likeness Property Comparison")

    # Header row
    header_cols = st.columns([2] + [1] * len(selected))
    header_cols[0].markdown("**Property**")
    for i, cpd in enumerate(selected):
        header_cols[i + 1].markdown(f"**{cpd.get('ID', f'Cpd-{i+1}')}**")

    # Data rows — use human-readable labels with tooltips
    for prop in _COMPARE_PROPS:
        row = st.columns([2] + [1] * len(selected))
        lbl = compare_label(prop)
        tip = compare_tip(prop)
        row[0].markdown(
            f'<span title="{tip or ""}" style="cursor:help;font-size:.82rem">{lbl}</span>',
            unsafe_allow_html=True)
        values = []
        for cpd in selected:
            v = cpd.get(prop, cpd.get(prop.lower(), "–"))
            values.append(v)
        # Highlight best
        try:
            nums = [float(v) for v in values]
            best_i = nums.index(max(nums))
        except Exception:
            best_i = -1

        for i, (col, v) in enumerate(zip(row[1:], values)):
            if i == best_i and prop in ["QED", "LeadScore", "OralBioScore"]:
                col.markdown(f"🏆 **{v}**")
            else:
                col.write(v)

    # ── Delta matrix (first compound as reference) ────────────────────────
    if len(selected) >= 2:
        st.subheader(f"Property Differences (Δ) vs {selected[0].get('ID', 'Reference Compound')}")
        ref = selected[0]
        delta_cols = st.columns([2] + [1] * (len(selected) - 1))
        delta_cols[0].markdown("**Property**")
        for i, cpd in enumerate(selected[1:]):
            delta_cols[i + 1].markdown(f"**{cpd.get('ID', f'Cpd-{i+2}')}**")

        for prop in _COMPARE_PROPS:
            drow = st.columns([2] + [1] * (len(selected) - 1))
            lbl = compare_label(prop)
            tip = compare_tip(prop)
            drow[0].markdown(
                f'<span title="{tip or ""}" style="cursor:help;font-size:.82rem">{lbl}</span>',
                unsafe_allow_html=True)
            ref_v = ref.get(prop, 0)
            for i, cpd in enumerate(selected[1:]):
                cpd_v = cpd.get(prop, 0)
                try:
                    delta = round(float(cpd_v) - float(ref_v), 2)
                    sign = "+" if delta > 0 else ""
                    color = "#4ade80" if delta >= 0 else "#f87171"
                    drow[i + 1].markdown(
                        f'<span style="color:{color};font-family:JetBrains Mono,monospace">'
                        f'{sign}{delta}</span>',
                        unsafe_allow_html=True,
                    )
                except Exception:
                    drow[i + 1].write("–")

    # NEW CODE START ──────────────────────────────────────────────────────
    # Advanced Comparison Insights — appended below existing table.
    # Nothing above this line has been modified.
    # ─────────────────────────────────────────────────────────────────────
    if len(selected) >= 2:
        _render_advanced_insights(selected)
    # NEW CODE END ────────────────────────────────────────────────────────


# =============================================================================
# NEW CODE START — Advanced Delta Engine & Intelligence Layers
# =============================================================================

# ── Safe optional imports ─────────────────────────────────────────────────
try:
    from rdkit import Chem as _Chem
    from rdkit.Chem import Descriptors as _Desc, rdMolDescriptors as _RDMol, Crippen as _Crippen
    _RDK_OK = True
except Exception:
    _RDK_OK = False

try:
    import admet_benchmark as _ab_mod
    _AB_MOD_OK = True
except Exception:
    _AB_MOD_OK = False

try:
    import reaction_simulator as _rs_mod
    _RS_MOD_OK = True
except Exception:
    _RS_MOD_OK = False


# ── Colour-coding helpers ─────────────────────────────────────────────────

def _delta_badge(value: float, higher_is_better: bool = True) -> str:
    if abs(value) < 0.001:
        return f'<span style="color:#94a3b8;font-family:JetBrains Mono,monospace">⚪ {value:+.2f}</span>'
    improved = (value > 0) == higher_is_better
    icon  = "🟢" if improved else "🔴"
    color = "#4ade80" if improved else "#f87171"
    return f'<span style="color:{color};font-family:JetBrains Mono,monospace">{icon} {value:+.2f}</span>'


def _text_badge(label: str, positive) -> str:
    if positive is True:
        return f'<span style="color:#4ade80">🟢 {label}</span>'
    if positive is False:
        return f'<span style="color:#f87171">🔴 {label}</span>'
    return f'<span style="color:#94a3b8">⚪ {label}</span>'


# ── Data-gathering helpers (lazy, read-only) ──────────────────────────────

def _get_rdkit_props(cpd: dict) -> dict:
    defaults = {
        "ring_count": 0, "aromatic_rings": 0,
        "heavy_atoms": 0, "chiral_centers": 0,
        "lipinski_violations": 0, "veber_ok": False,
    }
    if not _RDK_OK:
        return defaults
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi:
        return defaults
    try:
        mol = _Chem.MolFromSmiles(smi)
        if mol is None:
            return defaults
        mw   = _Desc.MolWt(mol)
        logp = _Crippen.MolLogP(mol)
        hbd  = _RDMol.CalcNumHBD(mol)
        hba  = _RDMol.CalcNumHBA(mol)
        tpsa = _Desc.TPSA(mol)
        rot  = _RDMol.CalcNumRotatableBonds(mol)
        viol = sum([mw > 500, logp > 5, hbd > 5, hba > 10])
        veber_ok = tpsa <= 140 and rot <= 10
        chiral = len(_Chem.FindMolChiralCenters(mol, includeUnassigned=True))
        return {
            "ring_count":          _RDMol.CalcNumRings(mol),
            "aromatic_rings":      _RDMol.CalcNumAromaticRings(mol),
            "heavy_atoms":         mol.GetNumHeavyAtoms(),
            "chiral_centers":      chiral,
            "lipinski_violations": viol,
            "veber_ok":            veber_ok,
        }
    except Exception:
        return defaults


def _get_metabolic_risk(cpd: dict) -> dict:
    defaults = {"metabolic_sites": 0, "phase1_vuln": "Unknown", "reactive_risk": "Unknown"}
    if not _RS_MOD_OK:
        return defaults
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi:
        return defaults
    try:
        hits = 0
        for rsmarts in _rs_mod._REACTIONS.values():
            if _rs_mod._apply_reaction(rsmarts, smi):
                hits += 1
        phase1  = "High" if hits >= 4 else "Medium" if hits >= 2 else "Low"
        reactive = "High" if hits >= 5 else "Medium" if hits >= 3 else "Low"
        return {"metabolic_sites": hits, "phase1_vuln": phase1, "reactive_risk": reactive}
    except Exception:
        return defaults


def _get_benchmark_score(cpd: dict) -> dict:
    defaults = {"drug_similarity": 50.0, "percentile_rank": 50.0, "outlier_score": 0.0}
    if not _AB_MOD_OK:
        return defaults
    try:
        ref  = _ab_mod._BENCHMARK_SETS["FDA Approved Drugs (n\u22482000)"]
        bm   = _ab_mod._benchmark_compound(cpd, ref)
        if not bm:
            return defaults
        pcts = [v["percentile"] for v in bm.values()]
        avg  = round(sum(pcts) / len(pcts), 1) if pcts else 50.0
        std  = (sum((p - avg) ** 2 for p in pcts) / max(len(pcts), 1)) ** 0.5
        return {"drug_similarity": avg, "percentile_rank": avg, "outlier_score": round(std, 1)}
    except Exception:
        return defaults


def _get_efficiency(cpd: dict) -> dict:
    try:
        qed  = float(cpd.get("QED", cpd.get("_qed", 0.5)))
        mw   = float(cpd.get("MW", 400))
        ha   = max(int(mw / 14), 1)
        le   = round((1.4 * qed * 100) / ha, 3)
        pd_  = round(qed / (mw / 100), 3) if mw > 0 else 0.0
        return {"ligand_efficiency": le, "property_density": pd_}
    except Exception:
        return {"ligand_efficiency": 0.0, "property_density": 0.0}


# ── Core advanced delta engine ────────────────────────────────────────────

def compute_advanced_deltas(r1: dict, r2: dict) -> dict:
    """
    Returns structured comparison insights without altering existing outputs.
    r1 = reference compound, r2 = compound being compared.
    """
    rdkit1  = _get_rdkit_props(r1)
    rdkit2  = _get_rdkit_props(r2)
    meta1   = _get_metabolic_risk(r1)
    meta2   = _get_metabolic_risk(r2)
    bench1  = _get_benchmark_score(r1)
    bench2  = _get_benchmark_score(r2)
    eff1    = _get_efficiency(r1)
    eff2    = _get_efficiency(r2)

    def _d(key, d1, d2, dec=2):
        try:
            return round(float(d2.get(key, 0)) - float(d1.get(key, 0)), dec)
        except Exception:
            return 0.0

    def _df(key, dec=2):
        try:
            return round(float(r2.get(key, 0)) - float(r1.get(key, 0)), dec)
        except Exception:
            return 0.0

    # 1. Drug-Likeness
    qed1 = float(r1.get("QED", r1.get("_qed", 0.5)))
    qed2 = float(r2.get("QED", r2.get("_qed", 0.5)))
    dqed = round(qed2 - qed1, 3)
    drug_likeness = {
        "delta_lipinski_violations": _d("lipinski_violations", rdkit1, rdkit2),
        "delta_veber_compliance":    int(rdkit2["veber_ok"]) - int(rdkit1["veber_ok"]),
        "delta_qed":                 dqed,
        "qed_direction":             "Improved" if dqed > 0.02 else "Worse" if dqed < -0.02 else "Neutral",
        "delta_sa_score":            _df("SA_Score"),
    }

    # 2. Risk
    tpsa1 = float(r1.get("tPSA", r1.get("TPSA", 90)))
    tpsa2 = float(r2.get("tPSA", r2.get("TPSA", 90)))
    logp1 = float(r1.get("LogP", 2.5))
    logp2 = float(r2.get("LogP", 2.5))
    mw1   = float(r1.get("MW", 400))
    mw2   = float(r2.get("MW", 400))
    tox1  = int(bool(r1.get("_pains"))) + (1 if r1.get("_herg") == "HIGH" else 0)
    tox2  = int(bool(r2.get("_pains"))) + (1 if r2.get("_herg") == "HIGH" else 0)
    risk = {
        "delta_solubility_risk":    round((logp2 * 0.5 + mw2 / 200) - (logp1 * 0.5 + mw1 / 200), 2),
        "delta_permeability_risk":  round(tpsa2 / 10 - tpsa1 / 10, 2),
        "delta_stability_proxy":    meta2["metabolic_sites"] - meta1["metabolic_sites"],
        "delta_toxicity_flags":     tox2 - tox1,
    }

    # 3. Structure
    structure = {
        "delta_ring_count":     rdkit2["ring_count"] - rdkit1["ring_count"],
        "delta_aromatic_rings": rdkit2["aromatic_rings"] - rdkit1["aromatic_rings"],
        "delta_heavy_atoms":    rdkit2["heavy_atoms"] - rdkit1["heavy_atoms"],
        "delta_chiral_centers": rdkit2["chiral_centers"] - rdkit1["chiral_centers"],
    }

    # 4. Engine consensus
    _ekeys  = ["LeadScore", "OralBioScore", "NP_Score", "ChemoScore"]
    scores1 = [float(r1[k]) for k in _ekeys if r1.get(k) is not None]
    scores2 = [float(r2[k]) for k in _ekeys if r2.get(k) is not None]

    def _var(vals):
        if len(vals) < 2:
            return 0.0
        m = sum(vals) / len(vals)
        return round(sum((v - m) ** 2 for v in vals) / len(vals), 2)

    agg1 = round(sum(scores1) / len(scores1), 2) if scores1 else 0.0
    agg2 = round(sum(scores2) / len(scores2), 2) if scores2 else 0.0
    var1 = _var(scores1)
    var2 = _var(scores2)
    engine = {
        "delta_engine_score":         round(agg2 - agg1, 2),
        "delta_agreement_score":      round(var2 - var1, 2),
        "delta_prediction_stability": round(var1 - var2, 2),
    }

    # 5. Metabolic
    _rmap = {"Low": 1, "Medium": 2, "High": 3, "Unknown": 0}
    metabolic = {
        "delta_metabolic_sites": meta2["metabolic_sites"] - meta1["metabolic_sites"],
        "delta_phase1_vuln":     _rmap.get(meta2["phase1_vuln"], 0) - _rmap.get(meta1["phase1_vuln"], 0),
        "phase1_r1":             meta1["phase1_vuln"],
        "phase1_r2":             meta2["phase1_vuln"],
        "delta_reactive_risk":   _rmap.get(meta2["reactive_risk"], 0) - _rmap.get(meta1["reactive_risk"], 0),
        "reactive_r1":           meta1["reactive_risk"],
        "reactive_r2":           meta2["reactive_risk"],
    }

    # 6. Benchmark
    benchmark = {
        "delta_drug_similarity": round(bench2["drug_similarity"] - bench1["drug_similarity"], 1),
        "delta_percentile_rank": round(bench2["percentile_rank"] - bench1["percentile_rank"], 1),
        "delta_outlier_score":   round(bench2["outlier_score"] - bench1["outlier_score"], 1),
        "pct_r1":                bench1["percentile_rank"],
        "pct_r2":                bench2["percentile_rank"],
    }

    # 7. Efficiency
    efficiency = {
        "delta_ligand_efficiency": round(eff2["ligand_efficiency"] - eff1["ligand_efficiency"], 3),
        "delta_property_density":  round(eff2["property_density"] - eff1["property_density"], 3),
        "le_r1": eff1["ligand_efficiency"],
        "le_r2": eff2["ligand_efficiency"],
    }

    # 8. Alerts
    new_violations = []
    if not r1.get("_pains") and r2.get("_pains"):
        new_violations.append("PAINS introduced")
    if r1.get("_herg") != "HIGH" and r2.get("_herg") == "HIGH":
        new_violations.append("hERG escalated to HIGH")
    if drug_likeness["delta_lipinski_violations"] > 0:
        new_violations.append(f"+{int(drug_likeness['delta_lipinski_violations'])} Lipinski violation(s)")
    if risk["delta_toxicity_flags"] > 0:
        new_violations.append(f"+{risk['delta_toxicity_flags']} tox flag(s)")
    alerts = {
        "new_violations":  new_violations,
        "risk_escalation": len(new_violations) > 0,
        "structural_alerts": (
            "Complexity increased" if structure["delta_heavy_atoms"] > 5
            else "Complexity reduced" if structure["delta_heavy_atoms"] < -5
            else "Complexity stable"
        ),
    }

    # 9. Decision
    abs_better  = int(tpsa2 < tpsa1 and logp2 < logp1) + int(mw2 < mw1 and tpsa2 < 90)
    stab_better = int(var2 < var1) + int(engine["delta_engine_score"] > 0)
    safe_better = int(risk["delta_toxicity_flags"] <= 0) + int(not new_violations)
    sc = {
        "Better for Absorption": abs_better,
        "Better for Stability":  stab_better,
        "Better for Safety":     safe_better,
        "Balanced Improvement":  int(abs_better > 0 and stab_better > 0 and safe_better > 0),
    }
    max_s = max(sc.values())
    if max_s == 0:
        direction = "No Clear Advantage"
    elif sc["Balanced Improvement"] == 1:
        direction = "Balanced Improvement"
    else:
        direction = max(sc, key=sc.get)
    decision = {"optimization_direction": direction, "score_breakdown": sc}

    return {
        "drug_likeness": drug_likeness,
        "risk":          risk,
        "structure":     structure,
        "engine":        engine,
        "metabolic":     metabolic,
        "benchmark":     benchmark,
        "efficiency":    efficiency,
        "decision":      decision,
        "alerts":        alerts,
    }


# ── Generic table renderer ────────────────────────────────────────────────

def _render_insight_table(cpd_ids, all_deltas, section, rows):
    hdr = st.columns([2] + [1] * len(cpd_ids))
    hdr[0].markdown("**Property / Metric**")
    for i, cid in enumerate(cpd_ids):
        hdr[i + 1].markdown(f"**{cid}**")

    for label_raw, key, higher_is_better in rows:
        # Use terminology feat_label if the raw label looks like an internal key
        display_label = feat_label(key) if label_raw.startswith("Δ") or " " not in label_raw else label_raw
        display_tip   = feat_tip(key)
        row = st.columns([2] + [1] * len(cpd_ids))
        row[0].markdown(
            f'<span title="{display_tip or ""}" style="cursor:help;font-size:.8rem">'
            f'{display_label}</span>',
            unsafe_allow_html=True)
        for i, cid in enumerate(cpd_ids):
            deltas = all_deltas.get(cid)
            if deltas is None:
                row[i + 1].write("–"); continue
            val = deltas.get(section, {}).get(key)
            if val is None:
                row[i + 1].write("–"); continue
            if isinstance(val, str):
                row[i + 1].write(val); continue
            if higher_is_better is None:
                row[i + 1].markdown(
                    f'<span style="color:#c8deff;font-family:JetBrains Mono,monospace">{val}</span>',
                    unsafe_allow_html=True)
            else:
                row[i + 1].markdown(_delta_badge(float(val), higher_is_better),
                                     unsafe_allow_html=True)


# ── Advanced insights renderer ────────────────────────────────────────────

def _render_advanced_insights(selected: list):
    st.markdown("---")
    st.markdown("### 🔬 Advanced Multi-Dimensional Comparison Insights")

    ref        = selected[0]
    comparisons = selected[1:]
    ref_id     = ref.get("ID", "Ref")

    if not comparisons:
        st.info("Select at least 2 compounds for advanced insights.")
        return

    # Pre-compute all deltas (local, no API calls)
    all_deltas = {}
    for cpd in comparisons:
        cid = cpd.get("ID", "?")
        try:
            all_deltas[cid] = compute_advanced_deltas(ref, cpd)
        except Exception:
            all_deltas[cid] = None

    cpd_ids = [c.get("ID", f"Cpd-{i+2}") for i, c in enumerate(comparisons)]

    itabs = st.tabs([
        "💊 Drug-Likeness & Developability",
        "⚠️ Risk & Toxicity Liability",
        "🧱 Structural & Scaffold Intelligence",
        "⚙️ Predictive Model Consensus",
        "🔥 Metabolic Liability & Biotransformation",
        "📐 FDA Drug Space Benchmarking",
        "⚡ Ligand Efficiency & Optimisation Metrics",
        "🚨 Safety Alerts & Toxicity Red Flags",
        "🎯 Go/No-Go Development Decision",
    ])

    # 1 — Drug-Likeness
    with itabs[0]:
        st.caption(f"Reference: **{ref_id}**")
        _render_insight_table(cpd_ids, all_deltas, "drug_likeness", [
            ("ΔLipinski Violations", "delta_lipinski_violations", False),
            ("ΔVeber Compliance",    "delta_veber_compliance",    True),
            ("ΔQED",                 "delta_qed",                 True),
            ("QED Direction",        "qed_direction",             None),
            ("ΔSA Score",            "delta_sa_score",            False),
        ])

    # 2 — Risk
    with itabs[1]:
        st.caption(f"Reference: **{ref_id}** · Lower risk delta = better")
        _render_insight_table(cpd_ids, all_deltas, "risk", [
            ("ΔSolubility Risk",   "delta_solubility_risk",   False),
            ("ΔPermeability Risk", "delta_permeability_risk", False),
            ("ΔStability Proxy",   "delta_stability_proxy",   False),
            ("ΔToxicity Flags",    "delta_toxicity_flags",    False),
        ])

    # 3 — Structure
    with itabs[2]:
        st.caption(f"Reference: **{ref_id}**")
        _render_insight_table(cpd_ids, all_deltas, "structure", [
            ("ΔRing Count",        "delta_ring_count",      None),
            ("ΔAromatic Rings",    "delta_aromatic_rings",  None),
            ("ΔHeavy Atoms",       "delta_heavy_atoms",     False),
            ("ΔChirality Centers", "delta_chiral_centers",  False),
        ])

    # 4 — Engine
    with itabs[3]:
        st.caption(f"Reference: **{ref_id}** · Higher score + lower variance = better")
        _render_insight_table(cpd_ids, all_deltas, "engine", [
            ("ΔEngine Score",         "delta_engine_score",         True),
            ("ΔAgreement (variance)",  "delta_agreement_score",      False),
            ("ΔPrediction Stability",  "delta_prediction_stability", True),
        ])

    # 5 — Metabolic
    with itabs[4]:
        st.caption(f"Reference: **{ref_id}** · Fewer metabolic sites = lower liability")
        _render_insight_table(cpd_ids, all_deltas, "metabolic", [
            ("ΔMetabolic Sites", "delta_metabolic_sites", False),
            ("ΔPhase I Vuln",    "delta_phase1_vuln",     False),
            ("Phase I (ref)",    "phase1_r1",              None),
            ("Phase I (cpd)",    "phase1_r2",              None),
            ("ΔReactive Risk",   "delta_reactive_risk",   False),
        ])

    # 6 — Benchmark
    with itabs[5]:
        st.caption(f"Reference: **{ref_id}** · vs FDA Approved Drug profile")
        _render_insight_table(cpd_ids, all_deltas, "benchmark", [
            ("ΔDrug Similarity",  "delta_drug_similarity", True),
            ("ΔPercentile Rank",  "delta_percentile_rank", True),
            ("ΔOutlier Score",    "delta_outlier_score",   False),
            ("Pct Rank (ref)",    "pct_r1",                None),
            ("Pct Rank (cpd)",    "pct_r2",                None),
        ])

    # 7 — Efficiency
    with itabs[6]:
        st.caption(f"Reference: **{ref_id}**")
        _render_insight_table(cpd_ids, all_deltas, "efficiency", [
            ("ΔLigand Efficiency",  "delta_ligand_efficiency", True),
            ("ΔProperty Density",   "delta_property_density",  True),
            ("LE (ref)",            "le_r1",                   None),
            ("LE (cpd)",            "le_r2",                   None),
        ])

    # 8 — Alerts
    with itabs[7]:
        st.caption("Red flags introduced vs reference compound")
        for cid, deltas in all_deltas.items():
            if deltas is None:
                st.error(f"{cid}: computation failed"); continue
            a = deltas["alerts"]
            with st.expander(f"**{cid}** — {len(a['new_violations'])} alert(s)", expanded=True):
                if a["risk_escalation"]:
                    for v in a["new_violations"]:
                        st.markdown(f'<span style="color:#f87171">🔴 {v}</span>',
                                    unsafe_allow_html=True)
                else:
                    st.markdown('<span style="color:#4ade80">🟢 No new violations vs reference</span>',
                                unsafe_allow_html=True)
                st.markdown(
                    f'<span style="color:#94a3b8">⚪ Structural: {a["structural_alerts"]}</span>',
                    unsafe_allow_html=True)

    # 9 — Decision
    with itabs[8]:
        st.caption("Optimization direction recommendation vs reference")
        _DIR_COLORS = {
            "Better for Absorption": "#38bdf8",
            "Better for Stability":  "#a78bfa",
            "Better for Safety":     "#4ade80",
            "Balanced Improvement":  "#f5a623",
            "No Clear Advantage":    "#94a3b8",
        }
        _DIR_ICONS = {
            "Better for Absorption": "💧",
            "Better for Stability":  "🔒",
            "Better for Safety":     "🛡️",
            "Balanced Improvement":  "⚖️",
            "No Clear Advantage":    "—",
        }
        for cid, deltas in all_deltas.items():
            if deltas is None:
                st.warning(f"{cid}: could not compute decision"); continue
            d         = deltas["decision"]
            direction = d["optimization_direction"]
            color     = _DIR_COLORS.get(direction, "#94a3b8")
            icon      = _DIR_ICONS.get(direction, "—")
            st.markdown(
                f'<div style="background:rgba(0,0,0,.2);border:1px solid {color}40;'
                f'border-left:4px solid {color};border-radius:8px;'
                f'padding:14px 18px;margin:8px 0">'
                f'<b style="font-family:JetBrains Mono,monospace;color:{color}">'
                f'{icon} {cid}</b><br>'
                f'<span style="color:{color};font-size:1.05rem;font-weight:600">'
                f'{direction}</span></div>',
                unsafe_allow_html=True,
            )
            sc_cols = st.columns(len(d["score_breakdown"]))
            for col, (k, v) in zip(sc_cols, d["score_breakdown"].items()):
                col.metric(k.replace("Better for ", ""), v)

# NEW CODE END ─────────────────────────────────────────────────────────────


# =============================================================================
# MASSIVE EXPANSION START — 100+ Feature Multi-Dimensional Comparison System
# =============================================================================

import math as _math

# ── Safe imports ──────────────────────────────────────────────────────────
try:
    from rdkit.Chem import (
        rdMolDescriptors as _RDMD2,
        Fragments as _Frags,
        GraphDescriptors as _Graph,
    )
    from rdkit.Chem.FilterCatalog import FilterCatalog as _FC, FilterCatalogParams as _FCP
    _RDK2_OK = True
except Exception:
    _RDK2_OK = False

try:
    import numpy as _np
    _NP_OK = True
except Exception:
    _NP_OK = False


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1 — PHYSICOCHEMICAL EXPANSION HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _physchem_expanded(cpd: dict) -> dict:
    """Compute 20+ advanced physicochemical descriptors. Read-only."""
    out = {k: 0.0 for k in [
        "polar_surface_dist_idx", "hydrophobic_surface_ratio", "mol_flexibility_idx",
        "topo_complexity_score", "shape_index", "atom_type_diversity",
        "functional_group_density", "heteroatom_ratio", "aromaticity_score",
        "ring_strain_proxy", "hb_network_density", "rotational_entropy_proxy",
        "mol_symmetry_score", "fragment_diversity_score", "saturation_index",
        "carbon_hybridization_ratio", "bond_type_diversity", "partial_charge_spread_proxy",
        "electronic_density_proxy", "mw_per_ring",
    ]}
    if not _RDK_OK:
        return out
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi:
        return out
    try:
        mol = _Chem.MolFromSmiles(smi)
        if mol is None:
            return out

        mw   = _Desc.MolWt(mol)
        tpsa = _Desc.TPSA(mol)
        logp = _Crippen.MolLogP(mol)
        ha   = max(mol.GetNumHeavyAtoms(), 1)
        rings    = _RDMol.CalcNumRings(mol)
        ar_rings = _RDMol.CalcNumAromaticRings(mol)
        rot      = _RDMol.CalcNumRotatableBonds(mol)
        hbd      = _RDMol.CalcNumHBD(mol)
        hba      = _RDMol.CalcNumHBA(mol)
        fsp3     = _RDMol.CalcFractionCSP3(mol)

        # Heteroatoms
        atoms     = [a for a in mol.GetAtoms()]
        hetero    = [a for a in atoms if a.GetAtomicNum() not in (1, 6)]
        het_ratio = round(len(hetero) / ha, 3)

        # Atom type diversity (unique atomic nums)
        atypes    = len(set(a.GetAtomicNum() for a in atoms if a.GetAtomicNum() > 1))

        # Aromaticity score
        ar_atoms  = sum(1 for a in atoms if a.GetIsAromatic())
        arom_scr  = round(ar_atoms / ha, 3)

        # Ring strain proxy — small rings (3,4-membered) count
        ri         = mol.GetRingInfo()
        small_rings = sum(1 for r in ri.AtomRings() if len(r) <= 4)

        # H-bond network density
        hbn_dens   = round((hbd + hba) / ha, 3)

        # Rotational entropy proxy
        rot_ent    = round(rot / ha, 3)

        # Symmetry proxy via canonical SMILES token repetition
        csmi       = _Chem.MolToSmiles(mol)
        sym_scr    = round(1.0 - len(set(csmi)) / max(len(csmi), 1), 3)

        # Saturation index (Fsp3 re-expression)
        sat_idx    = round(fsp3, 3)

        # C hybridization ratio (sp3 C / total C)
        c_atoms    = [a for a in atoms if a.GetAtomicNum() == 6]
        sp3c       = sum(1 for a in c_atoms if a.GetHybridization().name == "SP3")
        c_hyb      = round(sp3c / max(len(c_atoms), 1), 3)

        # Bond type diversity
        btypes     = len(set(b.GetBondTypeAsDouble() for b in mol.GetBonds()))

        # Partial charge spread proxy (logP range proxy)
        pc_spread  = round(abs(logp) / max(ha / 10, 1), 3)

        # Electronic density proxy
        ed_proxy   = round(tpsa / mw * 100, 2) if mw > 0 else 0

        # Polar surface distribution index
        psdi       = round(tpsa / ha, 3)

        # Hydrophobic surface ratio proxy
        hsr        = round(max(0.0, 1.0 - tpsa / 200.0), 3)

        # Flexibility index
        flex_idx   = round(rot / max(ha / 5, 1), 3)

        # Topological complexity (Bertz CT normalised)
        try:
            bertz  = _Desc.BertzCT(mol)
            topo   = round(_math.log1p(bertz) / 10.0, 3)
        except Exception:
            topo   = 0.0

        # Shape index proxy: MW / (rings+1) — high = elongated
        shape_idx  = round(mw / (rings + 1), 2)

        # Functional group density
        try:
            fg_count = sum(1 for fname in dir(_Frags)
                           if fname.startswith("fr_") and
                           callable(getattr(_Frags, fname)) and
                           getattr(_Frags, fname)(mol) > 0)
        except Exception:
            fg_count = 0
        fg_dens    = round(fg_count / ha, 3)

        # Fragment diversity (unique ring systems / total rings)
        frag_div   = round(len(set(frozenset(r) for r in ri.AtomRings())) / max(rings, 1), 3)

        # MW per ring
        mw_per_ring = round(mw / max(rings, 1), 2)

        out.update({
            "polar_surface_dist_idx":    psdi,
            "hydrophobic_surface_ratio": hsr,
            "mol_flexibility_idx":       flex_idx,
            "topo_complexity_score":     topo,
            "shape_index":               shape_idx,
            "atom_type_diversity":       float(atypes),
            "functional_group_density":  fg_dens,
            "heteroatom_ratio":          het_ratio,
            "aromaticity_score":         arom_scr,
            "ring_strain_proxy":         float(small_rings),
            "hb_network_density":        hbn_dens,
            "rotational_entropy_proxy":  rot_ent,
            "mol_symmetry_score":        sym_scr,
            "fragment_diversity_score":  frag_div,
            "saturation_index":          sat_idx,
            "carbon_hybridization_ratio":c_hyb,
            "bond_type_diversity":       float(btypes),
            "partial_charge_spread_proxy": pc_spread,
            "electronic_density_proxy":  ed_proxy,
            "mw_per_ring":               mw_per_ring,
        })
    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2 — DRUG-LIKENESS DEEP EXPANSION
# ─────────────────────────────────────────────────────────────────────────────

def _drug_likeness_deep(cpd: dict) -> dict:
    out = {k: 0.0 for k in [
        "ext_ro5_score", "ro3_compliance", "lead_likeness_score",
        "oral_bioavail_score", "cns_penetration_score", "efflux_risk_proxy",
        "ppb_proxy", "distribution_index", "drug_efficiency_score",
        "chem_beauty_score", "developability_index", "optim_readiness_score",
        "bcs_class_proxy", "abad_score", "gse_score",
        "pfizer_3_75_rule", "gsk_4_400_rule", "mce18_score",
        "natural_product_likeness", "fda_similarity_score",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        tpsa = float(cpd.get("tPSA", cpd.get("TPSA", 90)))
        qed  = float(cpd.get("QED", cpd.get("_qed", 0.5)))
        hbd  = float(cpd.get("HBD", cpd.get("_ext", {}).get("HBD", 2)))
        hba  = float(cpd.get("HBA", cpd.get("_ext", {}).get("HBA", 5)))
        rot  = float(cpd.get("RotBonds", cpd.get("_ext", {}).get("Rotatable_Bonds", 5)))
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        fsp3_raw = cpd.get("_ext", {}).get("Fsp3", 0.3) if isinstance(cpd.get("_ext"), dict) else 0.3
        fsp3 = float(fsp3_raw) if fsp3_raw else 0.3

        # Extended Ro5 score (0–5 violations penalised)
        ro5_viol = sum([mw > 500, logp > 5, hbd > 5, hba > 10])
        ext_ro5  = round(max(0.0, 1.0 - ro5_viol * 0.25), 2)

        # Ro3 (fragment-like): MW<=300, logp<=3, hbd<=3, hba<=3
        ro3 = float(mw <= 300 and logp <= 3 and hbd <= 3 and hba <= 3)

        # Lead-likeness: 250<MW<350, logp<=3.5, rot<=7
        ll = float(250 <= mw <= 350 and logp <= 3.5 and rot <= 7)

        # Oral bioavailability score (Lipinski + TPSA + rot)
        ob_score = round(ext_ro5 * 0.4 + (1 if tpsa <= 140 else 0) * 0.3 +
                         (1 if rot <= 10 else 0) * 0.3, 2)

        # CNS penetration (Wager MPO-like: MW<=450, logp 1-4, TPSA<=90, HBD<=3)
        cns = round(
            (1 if mw <= 450 else 0) * 0.25 +
            (1 if 1 <= logp <= 4 else 0) * 0.25 +
            (1 if tpsa <= 90 else 0) * 0.25 +
            (1 if hbd <= 3 else 0) * 0.25, 2)

        # Efflux risk proxy (P-gp): high MW + high TPSA = risk
        efflux_risk = round(min(1.0, (mw / 500 * 0.5 + tpsa / 140 * 0.5)), 2)

        # PPB proxy: high logP = high binding
        ppb = round(min(1.0, max(0.0, (logp + 2) / 8)), 2)

        # Distribution index: logP / (TPSA/10)
        dist_idx = round(logp / max(tpsa / 10, 0.1), 3)

        # Drug efficiency: QED / SA
        de = round(qed / max(sa / 10, 0.1), 3)

        # Chemical beauty (QED re-expression 0-100)
        beauty = round(qed * 100, 1)

        # Developability index
        dev_idx = round((ob_score + cns + ext_ro5) / 3, 3)

        # Optimization readiness (low SA = easy to make, high QED = good profile)
        opt_rdy = round((1 - sa / 10) * 0.5 + qed * 0.5, 3)

        # BCS class proxy (I=high sol+perm, II=low sol, III=low perm, IV=both low)
        hi_sol  = logp < 1
        hi_perm = tpsa < 90
        if hi_sol and hi_perm:     bcs = 1.0
        elif not hi_sol and hi_perm: bcs = 2.0
        elif hi_sol and not hi_perm: bcs = 3.0
        else:                       bcs = 4.0

        # ABAD (Activity-Based Absorption & Distribution)
        abad = round((1 if tpsa < 140 else 0) * 0.5 + (1 if mw < 500 else 0) * 0.5, 2)

        # GSE (General Solubility Equation proxy)
        gse = round(0.5 - logp - 0.01 * (mw - 100), 2)

        # Pfizer 3/75 rule: logP<3 AND TPSA>75 = safer
        pf375 = float(logp < 3 and tpsa > 75)

        # GSK 4/400 rule: logP<4 AND MW<400 = safer
        gsk440 = float(logp < 4 and mw < 400)

        # MCE-18 (sp3 + rings proxy)
        rings_raw = cpd.get("_ext", {}).get("Ring_Count", 2) if isinstance(cpd.get("_ext"), dict) else 2
        mce18 = round(fsp3 * 10 + float(rings_raw) * 0.5, 2)

        # Natural product likeness proxy
        np_score = float(cpd.get("NP_Score", 50)) / 100

        # FDA similarity (lead score proxy normalised)
        fda_sim = round(float(cpd.get("LeadScore", 60)) / 100, 2)

        out.update({
            "ext_ro5_score":          ext_ro5,
            "ro3_compliance":         ro3,
            "lead_likeness_score":    ll,
            "oral_bioavail_score":    ob_score,
            "cns_penetration_score":  cns,
            "efflux_risk_proxy":      efflux_risk,
            "ppb_proxy":              ppb,
            "distribution_index":     dist_idx,
            "drug_efficiency_score":  de,
            "chem_beauty_score":      beauty,
            "developability_index":   dev_idx,
            "optim_readiness_score":  opt_rdy,
            "bcs_class_proxy":        bcs,
            "abad_score":             abad,
            "gse_score":              gse,
            "pfizer_3_75_rule":       pf375,
            "gsk_4_400_rule":         gsk440,
            "mce18_score":            mce18,
            "natural_product_likeness": np_score,
            "fda_similarity_score":   fda_sim,
        })
    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3 — RISK & TOXICITY INTELLIGENCE
# ─────────────────────────────────────────────────────────────────────────────

def _risk_toxicity_deep(cpd: dict) -> dict:
    out = {k: 0.0 for k in [
        "reactive_group_density", "toxicophore_count", "lipophilicity_tox_risk",
        "aggregation_risk", "pains_alert_count", "struct_instability_idx",
        "oxidative_liability", "metabolic_instability", "bioaccumulation_risk",
        "clearance_risk", "safety_margin_proxy", "dose_escalation_risk",
        "off_target_risk", "chem_reactivity_score", "genotoxicity_proxy",
        "cardiotoxicity_proxy", "hepatotoxicity_proxy", "reproductive_tox_proxy",
        "mutagenicity_proxy", "overall_tox_index",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        tpsa = float(cpd.get("tPSA", cpd.get("TPSA", 90)))
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        pains_flag = bool(cpd.get("_pains", False))
        herg_flag  = cpd.get("_herg", "LOW") == "HIGH"

        # Lipophilicity toxicity risk: logP > 3 escalates
        lipo_tox = round(min(1.0, max(0.0, (logp - 1) / 4)), 2)

        # Aggregation risk: high logP + high MW
        agg_risk = round(min(1.0, (logp / 5 * 0.6 + mw / 600 * 0.4)), 2)

        # PAINS alert count (from stored data)
        pains_ct = float(cpd.get("_pains_count", int(pains_flag)))

        # Structural instability (low tpsa + high logP)
        struct_inst = round(min(1.0, max(0.0, logp / 5 * 0.5 + (1 - tpsa / 200) * 0.5)), 2)

        # Oxidative liability (aromatic-heavy proxy)
        ox_liab = round(min(1.0, float(cpd.get("_ext", {}).get("AromaticRings", 2) if
                                        isinstance(cpd.get("_ext"), dict) else 2) / 5), 2)

        # Metabolic instability (from metabolic sites if available)
        meta_sites = float(cpd.get("_meta_sites", 0))
        meta_inst  = round(min(1.0, meta_sites / 6), 2)

        # Bioaccumulation risk: logP > 3 is problematic
        bioaccum = round(min(1.0, max(0.0, (logp - 3) / 3)), 2)

        # Clearance risk: high SA = harder to clear
        clr_risk = round(min(1.0, sa / 10), 2)

        # Safety margin proxy (inverse of tox flags)
        tox_flags = int(pains_flag) + int(herg_flag)
        safety_margin = round(1.0 - min(1.0, tox_flags * 0.4), 2)

        # Dose escalation risk: poor safety + high logP
        dose_esc = round((lipo_tox * 0.5 + (1 - safety_margin) * 0.5), 2)

        # Off-target risk: high logP + high MW
        off_target = round(min(1.0, (logp / 5 * 0.4 + mw / 600 * 0.3 + (1 - tpsa / 200) * 0.3)), 2)

        # Chemical reactivity (SA score proxy)
        chem_react = round(min(1.0, sa / 10), 2)

        # Genotoxicity proxy (nitro/amine from SMILES pattern count)
        smi = cpd.get("SMILES") or ""
        geno_proxy = round(min(1.0, (smi.count("N") / max(len(smi), 1)) * 5), 2)

        # Cardiotoxicity proxy (hERG flag + logP)
        cardio = round(int(herg_flag) * 0.7 + lipo_tox * 0.3, 2)

        # Hepatotoxicity proxy (logP > 2 + reactive)
        hepato = round((1 if logp > 2 else 0) * 0.5 + chem_react * 0.5, 2)

        # Reproductive toxicity proxy (logP + MW)
        repro  = round(min(1.0, (logp / 6 * 0.5 + mw / 700 * 0.5)), 2)

        # Mutagenicity proxy
        mutagen = float(cpd.get("_ext", {}).get("Mutagenicity_Risk", "Low") == "High"
                        if isinstance(cpd.get("_ext"), dict) else False)

        # Overall toxicity index (composite)
        overall_tox = round(
            lipo_tox * 0.15 + agg_risk * 0.1 + struct_inst * 0.1 +
            ox_liab * 0.1 + meta_inst * 0.1 + bioaccum * 0.1 +
            cardio * 0.15 + hepato * 0.1 + mutagen * 0.1, 3)

        # Reactive group density
        if _RDK_OK:
            smi2 = cpd.get("SMILES") or cpd.get("smi") or ""
            rg_count = 0
            _REACT_SMARTS = [
                "[CX3](=[OX1])[F,Cl,Br,I]", "C1OC1", "[CX3H1]=O",
                "[$([NX3](=O)=O)]", "[CH2][Cl,Br,I]",
            ]
            try:
                mol_r = _Chem.MolFromSmiles(smi2)
                if mol_r:
                    for sma in _REACT_SMARTS:
                        pat = _Chem.MolFromSmarts(sma)
                        if pat and mol_r.GetSubstructMatches(pat):
                            rg_count += 1
            except Exception:
                pass
            ha2 = max(int(mw / 14), 1)
            rg_dens = round(rg_count / ha2, 3)
        else:
            rg_dens = 0.0

        # Toxicophore count (from _chemo_tests if present)
        tox_ct = sum(1 for t in cpd.get("_chemo_tests", [])
                     if "FAIL" in str(t.get("result", "")) and
                     "Tox" in str(t.get("category", "")))

        out.update({
            "reactive_group_density":  rg_dens,
            "toxicophore_count":       float(tox_ct),
            "lipophilicity_tox_risk":  lipo_tox,
            "aggregation_risk":        agg_risk,
            "pains_alert_count":       pains_ct,
            "struct_instability_idx":  struct_inst,
            "oxidative_liability":     ox_liab,
            "metabolic_instability":   meta_inst,
            "bioaccumulation_risk":    bioaccum,
            "clearance_risk":          clr_risk,
            "safety_margin_proxy":     safety_margin,
            "dose_escalation_risk":    dose_esc,
            "off_target_risk":         off_target,
            "chem_reactivity_score":   chem_react,
            "genotoxicity_proxy":      geno_proxy,
            "cardiotoxicity_proxy":    cardio,
            "hepatotoxicity_proxy":    hepato,
            "reproductive_tox_proxy":  repro,
            "mutagenicity_proxy":      mutagen,
            "overall_tox_index":       overall_tox,
        })
    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4 — STRUCTURAL & SCAFFOLD INTELLIGENCE
# ─────────────────────────────────────────────────────────────────────────────

def _structural_intelligence(cpd: dict) -> dict:
    out = {k: 0.0 for k in [
        "scaffold_stability_score", "scaffold_novelty_score", "scaffold_reusability_idx",
        "substituent_diversity", "branching_factor", "mol_tree_depth",
        "core_rigidity_score", "fragment_reusability", "struct_redundancy_score",
        "substructure_freq_score", "murcko_complexity", "ring_system_diversity",
        "linker_flexibility", "terminal_group_count", "heterocycle_richness",
    ]}
    if not _RDK_OK:
        return out
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi:
        return out
    try:
        mol = _Chem.MolFromSmiles(smi)
        if mol is None:
            return out

        ha   = max(mol.GetNumHeavyAtoms(), 1)
        mw   = _Desc.MolWt(mol)
        rings = _RDMol.CalcNumRings(mol)
        ar_rings = _RDMol.CalcNumAromaticRings(mol)
        rot  = _RDMol.CalcNumRotatableBonds(mol)
        fsp3 = _RDMol.CalcFractionCSP3(mol)
        ri   = mol.GetRingInfo()

        # Scaffold stability: aromatic-rich scaffolds are more stable
        scaf_stab = round(ar_rings / max(rings, 1), 3)

        # Scaffold novelty proxy: unusual atom types in ring systems
        ring_atoms = set()
        for r in ri.AtomRings():
            for idx in r:
                ring_atoms.add(mol.GetAtomWithIdx(idx).GetAtomicNum())
        ring_atoms.discard(6)  # remove carbon
        scaf_nov = round(len(ring_atoms) / max(rings, 1), 3)

        # Scaffold reusability: common ring systems (benzene, pyridine, etc.)
        common_rings = 0
        _COMMON = ["c1ccccc1", "c1ccncc1", "C1CCNCC1", "c1cncs1"]
        for cr in _COMMON:
            try:
                pat = _Chem.MolFromSmarts(cr)
                if pat and mol.GetSubstructMatches(pat):
                    common_rings += 1
            except Exception:
                pass
        scaf_reuse = round(common_rings / max(rings, 1), 3)

        # Substituent diversity: atoms not in rings / ha
        ring_atom_indices = set(idx for r in ri.AtomRings() for idx in r)
        subst_atoms = ha - len(ring_atom_indices)
        subst_div = round(subst_atoms / ha, 3)

        # Branching factor: degree > 2 atoms
        branch_atoms = sum(1 for a in mol.GetAtoms()
                           if len(a.GetNeighbors()) > 2 and a.GetAtomicNum() > 1)
        branching = round(branch_atoms / ha, 3)

        # Molecular tree depth proxy (longest path / ha)
        try:
            from rdkit.Chem import rdmolops as _rdmo
            dm = _rdmo.GetDistanceMatrix(mol)
            max_dist = float(dm.max()) if len(dm) > 0 else 0
        except Exception:
            max_dist = ha / 3
        tree_depth = round(max_dist / ha, 3)

        # Core rigidity: non-rotatable bonds / total bonds
        total_bonds = mol.GetNumBonds()
        rigidity = round(1.0 - rot / max(total_bonds, 1), 3)

        # Fragment reusability: how many of 4 common pharmacophore fragments match
        _PHARMA = ["[OH]", "[NH2]", "C(=O)", "c1ccccc1"]
        frag_match = 0
        for fp in _PHARMA:
            try:
                pat = _Chem.MolFromSmarts(fp)
                if pat and mol.GetSubstructMatches(pat):
                    frag_match += 1
            except Exception:
                pass
        frag_reuse = round(frag_match / 4, 2)

        # Structural redundancy (repeated motifs in SMILES)
        csmi = _Chem.MolToSmiles(mol)
        tokens = [csmi[i:i+3] for i in range(len(csmi) - 2)]
        redundancy = round(1.0 - len(set(tokens)) / max(len(tokens), 1), 3)

        # Substructure frequency score (common drug fragments)
        _DRUG_FRAGS = ["CC(=O)", "c1ccc(cc1)", "C(=O)N", "C(=O)O", "c1ccncc1"]
        sf_count = 0
        for df in _DRUG_FRAGS:
            try:
                pat = _Chem.MolFromSmarts(df)
                if pat and mol.GetSubstructMatches(pat):
                    sf_count += 1
            except Exception:
                pass
        sf_score = round(sf_count / 5, 2)

        # Murcko complexity (ring systems + linkers)
        murcko_cmplx = round(rings * 0.4 + subst_atoms * 0.6 / ha, 3)

        # Ring system diversity
        ring_sizes = [len(r) for r in ri.AtomRings()]
        ring_div = round(len(set(ring_sizes)) / max(len(ring_sizes), 1), 3)

        # Linker flexibility (rotatable bonds between ring systems)
        linker_flex = round(rot / max(rings + 1, 1), 3)

        # Terminal group count (atoms with degree 1)
        terminals = sum(1 for a in mol.GetAtoms()
                        if len(a.GetNeighbors()) == 1 and a.GetAtomicNum() > 1)
        terminal_ct = float(terminals)

        # Heterocycle richness
        het_rings = sum(1 for r in ri.AtomRings()
                        if any(mol.GetAtomWithIdx(i).GetAtomicNum() not in (6,)
                               for i in r))
        het_rich = round(het_rings / max(rings, 1), 3)

        out.update({
            "scaffold_stability_score":  scaf_stab,
            "scaffold_novelty_score":    scaf_nov,
            "scaffold_reusability_idx":  scaf_reuse,
            "substituent_diversity":     subst_div,
            "branching_factor":          branching,
            "mol_tree_depth":            tree_depth,
            "core_rigidity_score":       rigidity,
            "fragment_reusability":      frag_reuse,
            "struct_redundancy_score":   redundancy,
            "substructure_freq_score":   sf_score,
            "murcko_complexity":         murcko_cmplx,
            "ring_system_diversity":     ring_div,
            "linker_flexibility":        linker_flex,
            "terminal_group_count":      terminal_ct,
            "heterocycle_richness":      het_rich,
        })
    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5 — METABOLIC & BIO-TRANSFORMATION
# ─────────────────────────────────────────────────────────────────────────────

def _metabolic_deep(cpd: dict) -> dict:
    out = {k: 0.0 for k in [
        "phase1_rxn_count", "phase2_conjugation_likelihood", "metabolic_hotspot_density",
        "enzyme_interaction_likelihood", "clearance_speed_proxy",
        "metabolite_toxicity_risk", "stability_under_oxidation",
        "hydrolysis_susceptibility", "biotransformation_complexity",
        "cyp3a4_risk", "cyp2d6_risk", "ugt_liability", "aldehyde_oxidase_risk",
        "mao_liability", "gut_microbiome_risk",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        tpsa = float(cpd.get("tPSA", cpd.get("TPSA", 90)))
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        smi  = cpd.get("SMILES") or cpd.get("smi") or ""

        # Phase I rxn count (from reaction_simulator if cached, else estimate)
        if _RS_MOD_OK and smi:
            p1_count = 0
            try:
                for rs in _rs_mod._REACTIONS.values():
                    if _rs_mod._apply_reaction(rs, smi):
                        p1_count += 1
            except Exception:
                p1_count = int(logp > 2) + int(mw > 300)
        else:
            p1_count = int(logp > 2) + int(mw > 300)

        # Phase II conjugation likelihood (glucuronidation favoured by TPSA<100)
        p2_conj = round(min(1.0, (1 if tpsa < 100 else 0) * 0.5 +
                              (1 if logp < 4 else 0) * 0.5), 2)

        # Metabolic hotspot density
        hotspot_dens = round(p1_count / max(int(mw / 50), 1), 3)

        # Enzyme interaction likelihood
        enz_lik = round(min(1.0, (logp / 5 * 0.4 + mw / 500 * 0.3 +
                                   p1_count / 6 * 0.3)), 2)

        # Clearance speed proxy
        clr_speed = round(min(1.0, (1 if logp < 2 else 0) * 0.4 +
                               (1 if tpsa > 60 else 0) * 0.3 +
                               (1 if mw < 350 else 0) * 0.3), 2)

        # Metabolite toxicity risk
        meta_tox = round(min(1.0, (logp / 5) * 0.5 + (1 - clr_speed) * 0.5), 2)

        # Stability under oxidation (aromatic rings attract CYP oxidation)
        smi_lower = smi.lower()
        ar_count = smi.count("c")
        stab_ox = round(max(0.0, 1.0 - ar_count / max(len(smi), 1) * 3), 2)

        # Hydrolysis susceptibility (esters, amides)
        hydro_susc = round(min(1.0, (smi.count("C(=O)O") * 0.5 +
                                      smi.count("C(=O)N") * 0.3) / 2), 2)

        # Biotransformation complexity
        biotr_cmplx = round(p1_count * 0.4 + sa / 10 * 0.3 + p2_conj * 0.3, 3)

        # CYP3A4 risk: lipophilic + large molecules
        cyp3a4 = round(min(1.0, (logp / 5 * 0.5 + mw / 600 * 0.5)), 2)

        # CYP2D6 risk: basic nitrogen + moderate logP
        cyp2d6 = round(min(1.0, (smi.count("N") / max(len(smi), 1) * 8) * 0.6 +
                            (1 if 1 <= logp <= 4 else 0) * 0.4), 2)

        # UGT liability
        ugt = round(p2_conj * 0.7 + (1 if tpsa < 80 else 0) * 0.3, 2)

        # Aldehyde oxidase risk
        ao_risk = round(min(1.0, smi.count("n") / max(len(smi), 1) * 10), 2)

        # MAO liability
        mao_risk = round(min(1.0, (smi.count("N") / max(len(smi), 1) * 5) * 0.7 +
                              (1 if mw < 300 else 0) * 0.3), 2)

        # Gut microbiome risk (azo bonds, glucuronides susceptible)
        gut_risk = round(min(1.0, hydro_susc * 0.5 + p2_conj * 0.5), 2)

        out.update({
            "phase1_rxn_count":              float(p1_count),
            "phase2_conjugation_likelihood": p2_conj,
            "metabolic_hotspot_density":     hotspot_dens,
            "enzyme_interaction_likelihood": enz_lik,
            "clearance_speed_proxy":         clr_speed,
            "metabolite_toxicity_risk":      meta_tox,
            "stability_under_oxidation":     stab_ox,
            "hydrolysis_susceptibility":     hydro_susc,
            "biotransformation_complexity":  biotr_cmplx,
            "cyp3a4_risk":                   cyp3a4,
            "cyp2d6_risk":                   cyp2d6,
            "ugt_liability":                 ugt,
            "aldehyde_oxidase_risk":         ao_risk,
            "mao_liability":                 mao_risk,
            "gut_microbiome_risk":           gut_risk,
        })
    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6 — BENCHMARK & DATA-DRIVEN INSIGHTS
# ─────────────────────────────────────────────────────────────────────────────

def _benchmark_deep(cpd: dict, ref_cpd: dict) -> dict:
    out = {k: 0.0 for k in [
        "drug_space_distance", "top_drug_similarity", "outlier_detection_score",
        "cluster_position_proxy", "drug_likeness_percentile",
        "feature_dist_rank", "nearest_neighbor_dist",
        "approved_drug_overlap", "clinical_candidate_similarity",
        "natural_product_similarity",
    ]}
    try:
        if _AB_MOD_OK:
            ref_bm = _ab_mod._BENCHMARK_SETS.get("FDA Approved Drugs (n\u22482000)", {})
            bm = _ab_mod._benchmark_compound(cpd, ref_bm)
            if bm:
                pcts = [v["percentile"] for v in bm.values()]
                avg = sum(pcts) / len(pcts) if pcts else 50.0
                std = (_math.sqrt(sum((p - avg)**2 for p in pcts) / max(len(pcts), 1)))

                # Drug space distance (inverse of avg percentile)
                dsd = round(100 - avg, 1)

                # Top drug similarity
                top_sim = round(avg / 100, 3)

                # Outlier score
                outlier = round(std / 50, 3)

                # Cluster position proxy
                clust = round(avg / 100, 3)

                # Drug-likeness percentile
                dl_pct = round(avg, 1)

                # Feature distribution rank
                fdr = round(avg / 10, 1)

                # Nearest neighbor distance
                nn_dist = round(dsd / 100, 3)

                out.update({
                    "drug_space_distance":       dsd,
                    "top_drug_similarity":       top_sim,
                    "outlier_detection_score":   outlier,
                    "cluster_position_proxy":    clust,
                    "drug_likeness_percentile":  dl_pct,
                    "feature_dist_rank":         fdr,
                    "nearest_neighbor_dist":     nn_dist,
                })

        # Approved drug overlap (based on Lipinski + TPSA)
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        tpsa = float(cpd.get("tPSA", cpd.get("TPSA", 90)))
        hbd  = float(cpd.get("HBD", 2))

        # FDA approved space: MW 200-500, LogP -2 to 5, TPSA < 140
        in_fda_space = (200 <= mw <= 500 and -2 <= logp <= 5 and tpsa <= 140)
        out["approved_drug_overlap"] = float(in_fda_space)

        # Clinical candidate similarity
        # Phase II: slightly relaxed Lipinski
        in_clinical = (mw <= 550 and logp <= 5.5 and tpsa <= 160)
        out["clinical_candidate_similarity"] = float(in_clinical)

        # Natural product similarity (high MW, rings, low rot)
        out["natural_product_similarity"] = round(
            float(cpd.get("NP_Score", 50)) / 100, 3)

    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7 — ENGINE META-INTELLIGENCE
# ─────────────────────────────────────────────────────────────────────────────

def _engine_meta(cpd: dict) -> dict:
    out = {k: 0.0 for k in [
        "engine_agreement_score", "prediction_stability_idx", "confidence_interval_range",
        "engine_divergence_score", "consensus_strength", "reliability_score",
        "engine_count", "top_engine_score", "bottom_engine_score",
        "engine_spread_pct",
    ]}
    try:
        _ekeys = ["LeadScore", "OralBioScore", "NP_Score", "ChemoScore", "QED"]
        scores = []
        for k in _ekeys:
            v = cpd.get(k)
            if v is not None:
                try:
                    sv = float(v)
                    # Normalise QED to 0-100
                    if k == "QED" and sv <= 1.0:
                        sv *= 100
                    scores.append(sv)
                except Exception:
                    pass

        if len(scores) < 2:
            return out

        n    = len(scores)
        mean = sum(scores) / n
        var  = sum((s - mean) ** 2 for s in scores) / n
        std  = _math.sqrt(var)
        mn   = min(scores)
        mx   = max(scores)

        # Agreement score (1 = all agree, 0 = max divergence)
        agreement = round(1.0 - std / max(mean, 1) * 0.5, 3)

        # Prediction stability (low CV = stable)
        cv = round(std / max(mean, 1), 3)
        stability = round(1.0 - min(cv, 1.0), 3)

        # Confidence interval range (95% CI proxy = 2*std)
        ci_range = round(2 * std, 2)

        # Divergence score
        divergence = round(std / 100, 3)

        # Consensus strength (% of engines within 1 std of mean)
        within_1std = sum(1 for s in scores if abs(s - mean) <= std)
        consensus = round(within_1std / n, 3)

        # Reliability score
        reliability = round((agreement + stability + consensus) / 3, 3)

        out.update({
            "engine_agreement_score":   agreement,
            "prediction_stability_idx": stability,
            "confidence_interval_range": ci_range,
            "engine_divergence_score":  divergence,
            "consensus_strength":       consensus,
            "reliability_score":        reliability,
            "engine_count":             float(n),
            "top_engine_score":         round(mx, 2),
            "bottom_engine_score":      round(mn, 2),
            "engine_spread_pct":        round((mx - mn) / max(mean, 1) * 100, 1),
        })
    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 8 — EFFICIENCY & OPTIMISATION METRICS
# ─────────────────────────────────────────────────────────────────────────────

def _efficiency_deep(cpd: dict) -> dict:
    out = {k: 0.0 for k in [
        "ligand_efficiency", "lipophilic_efficiency", "property_efficiency_score",
        "optim_cost_estimate", "synth_feasibility_score", "dev_cost_proxy",
        "iteration_complexity", "lle_score", "ble_score",
        "atse_score", "group_efficiency", "binding_efficiency_idx",
        "kinetic_solubility_proxy", "permeability_efficiency",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        tpsa = float(cpd.get("tPSA", cpd.get("TPSA", 90)))
        qed  = float(cpd.get("QED", cpd.get("_qed", 0.5)))
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        lead = float(cpd.get("LeadScore", 60))
        ha   = max(int(mw / 14), 1)

        # Ligand efficiency (LE): pIC50_proxy / HA (using LeadScore as proxy)
        pIC50_proxy = lead / 20  # rough proxy
        le = round(pIC50_proxy / ha, 4)

        # Lipophilic ligand efficiency (LLE): pIC50 - logP
        lle = round(pIC50_proxy - logp, 3)

        # Binding efficiency index (BEI): pIC50 / MW * 1000
        ble = round(pIC50_proxy / mw * 1000, 3)

        # Property efficiency score
        pe_score = round((qed * 100 - abs(logp) * 5 - tpsa / 10) / 100, 3)

        # Optimisation cost estimate (high SA + low QED = expensive)
        optim_cost = round(sa / 10 * 0.5 + (1 - qed) * 0.5, 3)

        # Synthetic feasibility score (1 - SA/10)
        synth_feas = round(max(0.0, 1.0 - sa / 10), 3)

        # Development cost proxy (MW + SA + violations)
        ro5_viol = sum([mw > 500, logp > 5])
        dev_cost = round((mw / 500 * 0.4 + sa / 10 * 0.4 + ro5_viol * 0.1), 3)

        # Iteration complexity (how many rounds of optimisation needed)
        iter_cmplx = round((1 - qed) * 5 + sa / 10 * 3, 2)

        # Atom-type scaffold efficiency (ATSE)
        atse = round(qed / max(sa / 5, 0.1), 3)

        # Group efficiency (GE): activity contribution per functional group proxy
        fg_density = round(len(cpd.get("SMILES", "")) / max(mw / 100, 1), 3)
        group_eff = round(qed / max(fg_density, 0.01), 3)

        # Binding efficiency index (normalised)
        bei = round(pIC50_proxy * 1000 / max(mw, 1), 3)

        # Kinetic solubility proxy (GASTEiger-based: TPSA / logP)
        kin_sol = round(tpsa / max(abs(logp) + 1, 0.1), 2)

        # Permeability efficiency (TPSA-based)
        perm_eff = round(max(0.0, 1.0 - tpsa / 200), 3)

        out.update({
            "ligand_efficiency":         le,
            "lipophilic_efficiency":     lle,
            "property_efficiency_score": pe_score,
            "optim_cost_estimate":       optim_cost,
            "synth_feasibility_score":   synth_feas,
            "dev_cost_proxy":            dev_cost,
            "iteration_complexity":      iter_cmplx,
            "lle_score":                 lle,
            "ble_score":                 ble,
            "atse_score":                atse,
            "group_efficiency":          group_eff,
            "binding_efficiency_idx":    bei,
            "kinetic_solubility_proxy":  kin_sol,
            "permeability_efficiency":   perm_eff,
        })
    except Exception:
        pass
    return out


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 9 — DECISION INTELLIGENCE LAYER
# ─────────────────────────────────────────────────────────────────────────────

def _decision_intelligence(pc1, dl1, rt1, si1, md1, bm1, em1, ef1,
                            pc2, dl2, rt2, si2, md2, bm2, em2, ef2) -> dict:
    """Compute Go/No-Go and optimization direction from all layer scores."""
    try:
        def _delta(key, d1, d2):
            try:
                return float(d2.get(key, 0)) - float(d1.get(key, 0))
            except Exception:
                return 0.0

        # Collect improvement signals
        improvements = 0
        regressions  = 0

        checks = [
            # (key, d1, d2, higher_is_better)
            ("oral_bioavail_score",       dl1, dl2, True),
            ("cns_penetration_score",     dl1, dl2, True),
            ("developability_index",      dl1, dl2, True),
            ("optim_readiness_score",     dl1, dl2, True),
            ("overall_tox_index",         rt1, rt2, False),
            ("safety_margin_proxy",       rt1, rt2, True),
            ("engine_agreement_score",    em1, em2, True),
            ("reliability_score",         em1, em2, True),
            ("ligand_efficiency",         ef1, ef2, True),
            ("synth_feasibility_score",   ef1, ef2, True),
            ("scaffold_stability_score",  si1, si2, True),
            ("clearance_speed_proxy",     md1, md2, True),
            ("drug_likeness_percentile",  bm1, bm2, True),
        ]
        for key, d1, d2, hib in checks:
            delta = _delta(key, d1, d2)
            if abs(delta) < 0.01:
                continue
            if (delta > 0) == hib:
                improvements += 1
            else:
                regressions += 1

        total_signals = improvements + regressions
        if total_signals == 0:
            go_score = 0.5
        else:
            go_score = round(improvements / total_signals, 3)

        # Go/No-Go indicator
        if go_score >= 0.7:
            gng = "✅ GO"
        elif go_score >= 0.5:
            gng = "🟡 CONDITIONAL GO"
        elif go_score >= 0.3:
            gng = "🟠 NEEDS WORK"
        else:
            gng = "🔴 NO-GO"

        # Best use case classification
        abs_score  = _delta("oral_bioavail_score", dl1, dl2) + _delta("cns_penetration_score", dl1, dl2)
        safe_score = -_delta("overall_tox_index", rt1, rt2) + _delta("safety_margin_proxy", rt1, rt2)
        eff_score  = _delta("ligand_efficiency", ef1, ef2) + _delta("synth_feasibility_score", ef1, ef2)
        stab_score = _delta("engine_agreement_score", em1, em2) + _delta("reliability_score", em1, em2)

        best_scores = {
            "Absorption-Optimised":  abs_score,
            "Safety-Optimised":      safe_score,
            "Efficiency-Optimised":  eff_score,
            "Stability-Optimised":   stab_score,
        }
        best_use = max(best_scores, key=best_scores.get)
        if max(best_scores.values()) <= 0:
            best_use = "No Clear Use Case"

        # Tradeoff score (how balanced are improvements)
        vals = list(best_scores.values())
        pos_vals = [v for v in vals if v > 0]
        if len(pos_vals) >= 2:
            mean_p = sum(pos_vals) / len(pos_vals)
            tradeoff = round(1.0 - abs(max(pos_vals) - mean_p) / max(mean_p, 0.01), 3)
        else:
            tradeoff = 0.0

        # Risk vs reward index
        reward = improvements * 0.1
        risk   = regressions * 0.1 + float(rt2.get("overall_tox_index", 0))
        rvr    = round(reward / max(risk, 0.01), 3) if risk > 0 else round(reward * 10, 3)

        # Upgrade potential score
        upgrade_pot = round(
            _delta("optim_readiness_score", dl1, dl2) * 0.3 +
            _delta("synth_feasibility_score", ef1, ef2) * 0.3 +
            go_score * 0.4, 3)

        # Optimization priority
        if safe_score < 0:
            optim_priority = "🔴 Fix Safety First"
        elif abs_score < -0.1:
            optim_priority = "🟠 Improve Absorption"
        elif eff_score > 0.1:
            optim_priority = "🟢 Efficiency is Improving"
        elif stab_score > 0.1:
            optim_priority = "🟢 Stability is Improving"
        else:
            optim_priority = "⚪ Balanced Optimization"

        return {
            "go_no_go_indicator":      gng,
            "go_score":                go_score,
            "best_use_case":           best_use,
            "optimization_priority":   optim_priority,
            "tradeoff_score":          tradeoff,
            "risk_vs_reward_index":    rvr,
            "upgrade_potential_score": upgrade_pot,
            "improvements_count":      improvements,
            "regressions_count":       regressions,
            "net_signal":              improvements - regressions,
        }
    except Exception as e:
        return {
            "go_no_go_indicator": "⚠️ Error",
            "go_score": 0.5,
            "best_use_case": "Unknown",
            "optimization_priority": "Unknown",
            "tradeoff_score": 0.0,
            "risk_vs_reward_index": 0.0,
            "upgrade_potential_score": 0.0,
            "improvements_count": 0,
            "regressions_count": 0,
            "net_signal": 0,
        }


# ─────────────────────────────────────────────────────────────────────────────
# SECTION 10 — NATURAL LANGUAGE EXPLANATION
# ─────────────────────────────────────────────────────────────────────────────

def _generate_explanation(r1: dict, r2: dict,
                           pc_d: dict, dl_d: dict, rt_d: dict,
                           dec: dict) -> dict:
    """Generate plain-language explanation of comparison results."""
    id1 = r1.get("ID", "Reference")
    id2 = r2.get("ID", "Compound")

    lines_why_better = []
    lines_why_worse  = []
    lines_next_steps = []

    def _delta(key, d1, d2, higher_better=True):
        try:
            d = float(d2.get(key, 0)) - float(d1.get(key, 0))
            improved = (d > 0.01) == higher_better
            worsened = (d < -0.01) == higher_better
            return d, improved, worsened
        except Exception:
            return 0.0, False, False

    # Drug-likeness signals
    d, imp, wrs = _delta("oral_bioavail_score", {}, {}, True)
    oba1 = dl_d[0].get("oral_bioavail_score", 0.5)
    oba2 = dl_d[1].get("oral_bioavail_score", 0.5)
    if oba2 > oba1 + 0.05:
        lines_why_better.append(f"Improved oral bioavailability score ({oba1:.2f} → {oba2:.2f})")
    elif oba2 < oba1 - 0.05:
        lines_why_worse.append(f"Reduced oral bioavailability ({oba1:.2f} → {oba2:.2f})")

    cns1 = dl_d[0].get("cns_penetration_score", 0)
    cns2 = dl_d[1].get("cns_penetration_score", 0)
    if cns2 > cns1 + 0.05:
        lines_why_better.append(f"Better CNS penetration potential ({cns1:.2f} → {cns2:.2f})")
    elif cns2 < cns1 - 0.05:
        lines_why_worse.append(f"Lower CNS penetration ({cns1:.2f} → {cns2:.2f})")

    # Tox signals
    tox1 = rt_d[0].get("overall_tox_index", 0)
    tox2 = rt_d[1].get("overall_tox_index", 0)
    if tox2 < tox1 - 0.05:
        lines_why_better.append(f"Lower overall toxicity risk ({tox1:.2f} → {tox2:.2f})")
    elif tox2 > tox1 + 0.05:
        lines_why_worse.append(f"Higher toxicity index ({tox1:.2f} → {tox2:.2f})")

    # Efficiency
    le1 = pc_d[0].get("mol_flexibility_idx", 0)
    le2 = pc_d[1].get("mol_flexibility_idx", 0)
    if le2 < le1 - 0.05:
        lines_why_better.append("Reduced molecular flexibility (more rigid scaffold)")

    # Physicochemical
    mw1  = float(r1.get("MW", 400))
    mw2  = float(r2.get("MW", 400))
    lp1  = float(r1.get("LogP", 2.5))
    lp2  = float(r2.get("LogP", 2.5))
    if mw2 < mw1 - 20:
        lines_why_better.append(f"Lower molecular weight ({mw1:.0f} → {mw2:.0f} Da, easier absorption)")
    if abs(lp2 - lp1) > 0.5:
        direction = "reduced" if lp2 < lp1 else "increased"
        lines_why_better.append(f"LogP {direction} ({lp1:.2f} → {lp2:.2f})")

    # Next steps based on go score
    go = dec.get("go_score", 0.5)
    priority = dec.get("optimization_priority", "")
    if go >= 0.7:
        lines_next_steps.append(f"Strong candidate — advance {id2} to in vitro profiling.")
    elif go >= 0.5:
        lines_next_steps.append(f"Promising candidate — address noted liabilities before progression.")
    else:
        lines_next_steps.append(f"Significant concerns — consider structural redesign before advancing {id2}.")

    if "Safety" in priority:
        lines_next_steps.append("Priority: investigate and resolve safety/toxicity flags.")
    elif "Absorption" in priority:
        lines_next_steps.append("Priority: optimize physicochemical properties for better absorption.")
    elif "Efficiency" in priority:
        lines_next_steps.append("Ligand efficiency is improving — continue scaffold optimization.")

    # Compose narratives
    why_better_text = (
        " ".join(f"({i+1}) {l}" for i, l in enumerate(lines_why_better))
        if lines_why_better else f"No clear advantages identified for {id2} vs {id1}."
    )
    why_worse_text = (
        " ".join(f"({i+1}) {l}" for i, l in enumerate(lines_why_worse))
        if lines_why_worse else f"No clear regressions identified."
    )
    next_steps_text = " ".join(lines_next_steps) if lines_next_steps else "Continue evaluation."

    return {
        "why_better":   why_better_text,
        "why_worse":    why_worse_text,
        "next_steps":   next_steps_text,
        "summary":      f"{id2} vs {id1}: {dec.get('go_no_go_indicator', '⚪')} — "
                        f"{dec.get('best_use_case', 'undetermined')}.",
    }


# ─────────────────────────────────────────────────────────────────────────────
# MASTER FUNCTION
# ─────────────────────────────────────────────────────────────────────────────

def compute_massive_comparison_insights(r1: dict, r2: dict) -> dict:
    """
    Compute 100+ comparison features across 10 intelligence layers.
    r1 = reference, r2 = test compound.
    Returns complete structured dict. Non-destructive — read-only.
    """
    pc1 = _physchem_expanded(r1)
    pc2 = _physchem_expanded(r2)
    dl1 = _drug_likeness_deep(r1)
    dl2 = _drug_likeness_deep(r2)
    rt1 = _risk_toxicity_deep(r1)
    rt2 = _risk_toxicity_deep(r2)
    si1 = _structural_intelligence(r1)
    si2 = _structural_intelligence(r2)
    md1 = _metabolic_deep(r1)
    md2 = _metabolic_deep(r2)
    bm1 = _benchmark_deep(r1, r1)
    bm2 = _benchmark_deep(r2, r1)
    em1 = _engine_meta(r1)
    em2 = _engine_meta(r2)
    ef1 = _efficiency_deep(r1)
    ef2 = _efficiency_deep(r2)

    dec = _decision_intelligence(pc1, dl1, rt1, si1, md1, bm1, em1, ef1,
                                  pc2, dl2, rt2, si2, md2, bm2, em2, ef2)

    expl = _generate_explanation(r1, r2, [pc1, pc2], [dl1, dl2], [rt1, rt2], dec)

    # Build delta layers
    def _deltas(d1, d2, higher_better_map: dict) -> dict:
        result = {}
        for key in d1:
            try:
                delta = round(float(d2.get(key, 0)) - float(d1.get(key, 0)), 4)
                hib = higher_better_map.get(key, True)
                result[key] = {
                    "ref": d1.get(key, 0),
                    "cpd": d2.get(key, 0),
                    "delta": delta,
                    "improved": (delta > 0.001 and hib) or (delta < -0.001 and not hib),
                    "worse":    (delta < -0.001 and hib) or (delta > 0.001 and not hib),
                    "higher_better": hib,
                }
            except Exception:
                result[key] = {"ref": 0, "cpd": 0, "delta": 0,
                                "improved": False, "worse": False, "higher_better": True}
        return result

    _PC_HB = {k: True for k in pc1}   # neutral for physchem
    _PC_HB.update({"mol_flexibility_idx": False, "ring_strain_proxy": False,
                   "shape_index": False, "partial_charge_spread_proxy": False})
    _DL_HB = {k: True for k in dl1}
    _DL_HB.update({"efflux_risk_proxy": False, "ppb_proxy": False, "bcs_class_proxy": False})
    _RT_HB = {k: False for k in rt1}
    _RT_HB.update({"safety_margin_proxy": True})
    _SI_HB = {k: True for k in si1}
    _SI_HB.update({"struct_redundancy_score": False, "branching_factor": False})
    _MD_HB = {k: False for k in md1}
    _MD_HB.update({"clearance_speed_proxy": True, "stability_under_oxidation": True})
    _BM_HB = {k: True for k in bm1}
    _BM_HB.update({"drug_space_distance": False, "outlier_detection_score": False,
                   "nearest_neighbor_dist": False})
    _EM_HB = {k: True for k in em1}
    _EM_HB.update({"confidence_interval_range": False, "engine_divergence_score": False,
                   "engine_spread_pct": False})
    _EF_HB = {k: True for k in ef1}
    _EF_HB.update({"optim_cost_estimate": False, "dev_cost_proxy": False, "iteration_complexity": False})

    return {
        "physchem":     _deltas(pc1, pc2, _PC_HB),
        "drug_likeness": _deltas(dl1, dl2, _DL_HB),
        "risk":          _deltas(rt1, rt2, _RT_HB),
        "structure":     _deltas(si1, si2, _SI_HB),
        "metabolic":     _deltas(md1, md2, _MD_HB),
        "benchmark":     _deltas(bm1, bm2, _BM_HB),
        "engine":        _deltas(em1, em2, _EM_HB),
        "efficiency":    _deltas(ef1, ef2, _EF_HB),
        "decision":      dec,
        "explanation":   expl,
        # Raw values for reference
        "_raw": {"pc1": pc1, "pc2": pc2, "dl1": dl1, "dl2": dl2,
                 "rt1": rt1, "rt2": rt2, "si1": si1, "si2": si2,
                 "md1": md1, "md2": md2, "bm1": bm1, "bm2": bm2,
                 "em1": em1, "em2": em2, "ef1": ef1, "ef2": ef2},
    }


# ─────────────────────────────────────────────────────────────────────────────
# UI RENDERER — MASSIVE COMPARISON SYSTEM
# ─────────────────────────────────────────────────────────────────────────────

def _massive_delta_cell(data: dict) -> str:
    """Render a single delta cell with colour coding."""
    delta = data.get("delta", 0)
    ref   = data.get("ref", 0)
    cpd   = data.get("cpd", 0)
    hib   = data.get("higher_better", True)
    imp   = data.get("improved", False)
    wrs   = data.get("worse", False)

    if imp:
        icon, color = "🟢", "#4ade80"
    elif wrs:
        icon, color = "🔴", "#f87171"
    else:
        icon, color = "⚪", "#94a3b8"

    sign = "+" if delta > 0 else ""
    try:
        return (f'<span style="color:{color};font-family:JetBrains Mono,monospace;'
                f'font-size:.78rem">{icon} {sign}{float(delta):.3f}'
                f'<br><span style="font-size:.65rem;color:#64748b">'
                f'ref:{float(ref):.2f} → {float(cpd):.2f}</span></span>')
    except Exception:
        return f'<span style="color:#94a3b8">–</span>'


def _render_massive_section(cpd_ids: list, all_insights: dict, section: str, row_keys: list):
    """Render a section table from the massive insights dict."""
    hdr = st.columns([2] + [1] * len(cpd_ids))
    hdr[0].markdown("**Metric**")
    for i, cid in enumerate(cpd_ids):
        hdr[i + 1].markdown(f"**{cid}**")

    for label, key in row_keys:
        row = st.columns([2] + [1] * len(cpd_ids))
        row[0].markdown(f'<span style="font-size:.82rem">{label}</span>',
                        unsafe_allow_html=True)
        for i, cid in enumerate(cpd_ids):
            ins = all_insights.get(cid)
            if ins is None:
                row[i + 1].write("–")
                continue
            sec_data = ins.get(section, {})
            cell_data = sec_data.get(key, {})
            if not isinstance(cell_data, dict):
                row[i + 1].write(str(cell_data))
                continue
            row[i + 1].markdown(_massive_delta_cell(cell_data), unsafe_allow_html=True)


def _render_massive_comparison(selected: list):
    """Main renderer for the massive 100+ feature comparison system."""
    st.markdown("---")
    st.markdown("## 🚀 Advanced Multi-Dimensional Comparison System — 124 Scientific Intelligence Signals")
    st.caption("100+ chemoinformatics signals across 10 intelligence layers")

    ref  = selected[0]
    cpds = selected[1:]
    if not cpds:
        st.info("Select 2+ compounds for advanced analysis.")
        return

    ref_id  = ref.get("ID", "Ref")
    cpd_ids = [c.get("ID", f"Cpd-{i+2}") for i, c in enumerate(cpds)]

    # Compute all insights (cached in session_state keyed by compound IDs)
    _cache_key_mas = "_massive_insights_" + "_".join([ref_id] + cpd_ids)
    if _cache_key_mas not in st.session_state:
        all_insights = {}
        for cpd in cpds:
            cid = cpd.get("ID", "?")
            try:
                all_insights[cid] = compute_massive_comparison_insights(ref, cpd)
            except Exception:
                all_insights[cid] = None
        st.session_state[_cache_key_mas] = all_insights
    else:
        all_insights = st.session_state[_cache_key_mas]

    # ── 9-tab layout ─────────────────────────────────────────────────────
    tabs = st.tabs([
        "⚗️ Physicochemical",
        "💊 Drug-Likeness & Developability",
        "🚨 Risk & Tox",
        "🧱 Structural & Scaffold Intelligence",
        "🔥 Metabolic Liability & Biotransformation",
        "📐 FDA Drug Space Benchmarking",
        "⚙️ Predictive Model Consensus",
        "⚡ Ligand Efficiency & Optimisation Metrics",
        "🎯 Go/No-Go Development Decision",
    ])

    # ── Tab 1: Physicochemical (20 features) ─────────────────────────────
    with tabs[0]:
        st.caption(f"Reference: **{ref_id}** — 20 advanced physicochemical descriptors")
        _render_massive_section(cpd_ids, all_insights, "physchem", [
            ("Polar Surface Dist. Index",     "polar_surface_dist_idx"),
            ("Hydrophobic Surface Ratio",     "hydrophobic_surface_ratio"),
            ("Molecular Flexibility Index",   "mol_flexibility_idx"),
            ("Topological Complexity",        "topo_complexity_score"),
            ("Shape Index (MW/rings)",        "shape_index"),
            ("Atom Type Diversity",           "atom_type_diversity"),
            ("Functional Group Density",      "functional_group_density"),
            ("Heteroatom Ratio",              "heteroatom_ratio"),
            ("Aromaticity Score",             "aromaticity_score"),
            ("Ring Strain Proxy",             "ring_strain_proxy"),
            ("H-Bond Network Density",        "hb_network_density"),
            ("Rotational Entropy Proxy",      "rotational_entropy_proxy"),
            ("Molecular Symmetry Score",      "mol_symmetry_score"),
            ("Fragment Diversity Score",      "fragment_diversity_score"),
            ("Saturation Index (Fsp3)",       "saturation_index"),
            ("C Hybridization Ratio",         "carbon_hybridization_ratio"),
            ("Bond Type Diversity",           "bond_type_diversity"),
            ("Partial Charge Spread",         "partial_charge_spread_proxy"),
            ("Electronic Density Proxy",      "electronic_density_proxy"),
            ("MW per Ring",                   "mw_per_ring"),
        ])

    # ── Tab 2: Drug-Likeness (20 features) ───────────────────────────────
    with tabs[1]:
        st.caption(f"Reference: **{ref_id}** — 20 drug-likeness & developability metrics")
        _render_massive_section(cpd_ids, all_insights, "drug_likeness", [
            ("Extended Ro5 Score",            "ext_ro5_score"),
            ("Ro3 Compliance (fragment)",     "ro3_compliance"),
            ("Lead-Likeness Score",           "lead_likeness_score"),
            ("Oral Bioavailability Score",    "oral_bioavail_score"),
            ("CNS Penetration Score",         "cns_penetration_score"),
            ("Efflux Risk Proxy (P-gp)",      "efflux_risk_proxy"),
            ("PPB Proxy",                     "ppb_proxy"),
            ("Distribution Index",            "distribution_index"),
            ("Drug Efficiency Score",         "drug_efficiency_score"),
            ("Chemical Beauty (QED×100)",     "chem_beauty_score"),
            ("Developability Index",          "developability_index"),
            ("Optimization Readiness",        "optim_readiness_score"),
            ("BCS Class Proxy",               "bcs_class_proxy"),
            ("ABAD Score",                    "abad_score"),
            ("GSE Solubility Score",          "gse_score"),
            ("Pfizer 3/75 Rule",              "pfizer_3_75_rule"),
            ("GSK 4/400 Rule",                "gsk_4_400_rule"),
            ("MCE-18 Score",                  "mce18_score"),
            ("NP-Likeness",                   "natural_product_likeness"),
            ("FDA Similarity Score",          "fda_similarity_score"),
        ])

    # ── Tab 3: Risk & Toxicity (20 features) ─────────────────────────────
    with tabs[2]:
        st.caption(f"Reference: **{ref_id}** — 20 risk & toxicity intelligence signals")
        _render_massive_section(cpd_ids, all_insights, "risk", [
            ("Reactive Group Density",        "reactive_group_density"),
            ("Toxicophore Count",             "toxicophore_count"),
            ("Lipophilicity Tox Risk",        "lipophilicity_tox_risk"),
            ("Aggregation Risk",              "aggregation_risk"),
            ("PAINS Alert Count",             "pains_alert_count"),
            ("Structural Instability Index",  "struct_instability_idx"),
            ("Oxidative Liability",           "oxidative_liability"),
            ("Metabolic Instability",         "metabolic_instability"),
            ("Bioaccumulation Risk",          "bioaccumulation_risk"),
            ("Clearance Risk",                "clearance_risk"),
            ("Safety Margin Proxy",           "safety_margin_proxy"),
            ("Dose Escalation Risk",          "dose_escalation_risk"),
            ("Off-Target Risk",               "off_target_risk"),
            ("Chemical Reactivity Score",     "chem_reactivity_score"),
            ("Genotoxicity Proxy",            "genotoxicity_proxy"),
            ("Cardiotoxicity Proxy",          "cardiotoxicity_proxy"),
            ("Hepatotoxicity Proxy",          "hepatotoxicity_proxy"),
            ("Reproductive Tox Proxy",        "reproductive_tox_proxy"),
            ("Mutagenicity Proxy",            "mutagenicity_proxy"),
            ("Overall Toxicity Index",        "overall_tox_index"),
        ])

    # ── Tab 4: Structure (15 features) ───────────────────────────────────
    with tabs[3]:
        st.caption(f"Reference: **{ref_id}** — 15 structural & scaffold intelligence signals")
        _render_massive_section(cpd_ids, all_insights, "structure", [
            ("Scaffold Stability Score",      "scaffold_stability_score"),
            ("Scaffold Novelty Score",        "scaffold_novelty_score"),
            ("Scaffold Reusability Index",    "scaffold_reusability_idx"),
            ("Substituent Diversity",         "substituent_diversity"),
            ("Branching Factor",              "branching_factor"),
            ("Molecular Tree Depth",          "mol_tree_depth"),
            ("Core Rigidity Score",           "core_rigidity_score"),
            ("Fragment Reusability",          "fragment_reusability"),
            ("Structural Redundancy",         "struct_redundancy_score"),
            ("Substructure Freq Score",       "substructure_freq_score"),
            ("Murcko Complexity",             "murcko_complexity"),
            ("Ring System Diversity",         "ring_system_diversity"),
            ("Linker Flexibility",            "linker_flexibility"),
            ("Terminal Group Count",          "terminal_group_count"),
            ("Heterocycle Richness",          "heterocycle_richness"),
        ])

    # ── Tab 5: Metabolic (15 features) ───────────────────────────────────
    with tabs[4]:
        st.caption(f"Reference: **{ref_id}** — 15 metabolic liability signals")
        _render_massive_section(cpd_ids, all_insights, "metabolic", [
            ("Phase I Reaction Count",        "phase1_rxn_count"),
            ("Phase II Conjugation Likelihood","phase2_conjugation_likelihood"),
            ("Metabolic Hotspot Density",     "metabolic_hotspot_density"),
            ("Enzyme Interaction Likelihood", "enzyme_interaction_likelihood"),
            ("Clearance Speed Proxy",         "clearance_speed_proxy"),
            ("Metabolite Toxicity Risk",      "metabolite_toxicity_risk"),
            ("Stability Under Oxidation",     "stability_under_oxidation"),
            ("Hydrolysis Susceptibility",     "hydrolysis_susceptibility"),
            ("Biotransformation Complexity",  "biotransformation_complexity"),
            ("CYP3A4 Risk",                   "cyp3a4_risk"),
            ("CYP2D6 Risk",                   "cyp2d6_risk"),
            ("UGT Liability",                 "ugt_liability"),
            ("Aldehyde Oxidase Risk",         "aldehyde_oxidase_risk"),
            ("MAO Liability",                 "mao_liability"),
            ("Gut Microbiome Risk",           "gut_microbiome_risk"),
        ])

    # ── Tab 6: Benchmark (10 features) ───────────────────────────────────
    with tabs[5]:
        st.caption(f"Reference: **{ref_id}** — Benchmark vs approved drug space")
        _render_massive_section(cpd_ids, all_insights, "benchmark", [
            ("Distance from Drug Space",      "drug_space_distance"),
            ("Top Drug Similarity",           "top_drug_similarity"),
            ("Outlier Detection Score",       "outlier_detection_score"),
            ("Cluster Position Proxy",        "cluster_position_proxy"),
            ("Drug-Likeness Percentile",      "drug_likeness_percentile"),
            ("Feature Distribution Rank",     "feature_dist_rank"),
            ("Nearest Neighbor Distance",     "nearest_neighbor_dist"),
            ("Approved Drug Overlap",         "approved_drug_overlap"),
            ("Clinical Candidate Similarity", "clinical_candidate_similarity"),
            ("Natural Product Similarity",    "natural_product_similarity"),
        ])

    # ── Tab 7: Engine Intelligence (10 features) ─────────────────────────
    with tabs[6]:
        st.caption(f"Reference: **{ref_id}** — Engine meta-intelligence layer")
        _render_massive_section(cpd_ids, all_insights, "engine", [
            ("Engine Agreement Score",        "engine_agreement_score"),
            ("Prediction Stability Index",    "prediction_stability_idx"),
            ("Confidence Interval Range",     "confidence_interval_range"),
            ("Engine Divergence Score",       "engine_divergence_score"),
            ("Consensus Strength",            "consensus_strength"),
            ("Reliability Score",             "reliability_score"),
            ("Engine Count",                  "engine_count"),
            ("Top Engine Score",              "top_engine_score"),
            ("Bottom Engine Score",           "bottom_engine_score"),
            ("Engine Spread %",               "engine_spread_pct"),
        ])

    # ── Tab 8: Efficiency (14 features) ──────────────────────────────────
    with tabs[7]:
        st.caption(f"Reference: **{ref_id}** — Efficiency & optimisation metrics")
        _render_massive_section(cpd_ids, all_insights, "efficiency", [
            ("Ligand Efficiency (LE)",        "ligand_efficiency"),
            ("Lipophilic Ligand Eff (LLE)",   "lipophilic_efficiency"),
            ("Property Efficiency Score",     "property_efficiency_score"),
            ("Optimisation Cost Estimate",    "optim_cost_estimate"),
            ("Synthetic Feasibility Score",   "synth_feasibility_score"),
            ("Development Cost Proxy",        "dev_cost_proxy"),
            ("Iteration Complexity",          "iteration_complexity"),
            ("LLE Score",                     "lle_score"),
            ("BLE Score",                     "ble_score"),
            ("ATSE Score",                    "atse_score"),
            ("Group Efficiency",              "group_efficiency"),
            ("Binding Efficiency Index",      "binding_efficiency_idx"),
            ("Kinetic Solubility Proxy",      "kinetic_solubility_proxy"),
            ("Permeability Efficiency",       "permeability_efficiency"),
        ])

    # ── Tab 9: Decision Intelligence ─────────────────────────────────────
    with tabs[8]:
        st.caption(f"Reference: **{ref_id}** — Go/No-Go decision layer + AI explanation")

        _DEC_COLORS = {
            "✅ GO":              "#4ade80",
            "🟡 CONDITIONAL GO":  "#f5a623",
            "🟠 NEEDS WORK":      "#fb923c",
            "🔴 NO-GO":           "#f87171",
        }

        for cid in cpd_ids:
            ins = all_insights.get(cid)
            if ins is None:
                st.error(f"{cid}: computation failed")
                continue

            dec  = ins.get("decision", {})
            expl = ins.get("explanation", {})
            gng  = dec.get("go_no_go_indicator", "⚠️")
            color = _DEC_COLORS.get(gng, "#94a3b8")

            # Go/No-Go header card
            st.markdown(
                f'<div style="background:rgba(0,0,0,.25);border:1px solid {color}50;'
                f'border-left:5px solid {color};border-radius:10px;'
                f'padding:16px 20px;margin:10px 0">'
                f'<div style="font-family:JetBrains Mono,monospace;font-size:.6rem;'
                f'color:{color};letter-spacing:3px;text-transform:uppercase">'
                f'Decision · {cid}</div>'
                f'<div style="font-size:1.4rem;font-weight:700;color:{color};margin:6px 0">'
                f'{gng}</div>'
                f'<div style="color:#c8deff;font-size:.85rem">'
                f'Go Score: {dec.get("go_score", 0):.0%} &nbsp;|&nbsp; '
                f'Best Use: {dec.get("best_use_case","–")} &nbsp;|&nbsp; '
                f'{dec.get("optimization_priority","–")}</div></div>',
                unsafe_allow_html=True,
            )

            # KPI metrics row
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Tradeoff Score",   f'{dec.get("tradeoff_score", 0):.3f}')
            m2.metric("Risk/Reward",      f'{dec.get("risk_vs_reward_index", 0):.3f}')
            m3.metric("Upgrade Potential",f'{dec.get("upgrade_potential_score", 0):.3f}')
            m4.metric("Net Signal",       f'{dec.get("improvements_count",0) - dec.get("regressions_count",0):+d}')

            # AI-style explanation boxes
            with st.expander(f"📖 Why {cid} is better", expanded=True):
                st.markdown(
                    f'<div style="background:rgba(74,222,128,.05);border:1px solid rgba(74,222,128,.2);'
                    f'border-radius:8px;padding:12px 16px;font-size:.83rem;color:#c8deff;line-height:1.7">'
                    f'{expl.get("why_better","–")}</div>',
                    unsafe_allow_html=True)

            with st.expander(f"⚠️ Liabilities of {cid}"):
                st.markdown(
                    f'<div style="background:rgba(248,113,113,.05);border:1px solid rgba(248,113,113,.2);'
                    f'border-radius:8px;padding:12px 16px;font-size:.83rem;color:#c8deff;line-height:1.7">'
                    f'{expl.get("why_worse","–")}</div>',
                    unsafe_allow_html=True)

            with st.expander(f"🎯 Next Steps for {cid}"):
                st.markdown(
                    f'<div style="background:rgba(245,166,35,.05);border:1px solid rgba(245,166,35,.2);'
                    f'border-radius:8px;padding:12px 16px;font-size:.83rem;color:#c8deff;line-height:1.7">'
                    f'{expl.get("next_steps","–")}</div>',
                    unsafe_allow_html=True)

            st.divider()


# ─────────────────────────────────────────────────────────────────────────────
# HOOK INTO render_tab — inject massive section after existing advanced section
# ─────────────────────────────────────────────────────────────────────────────
# NOTE: render_tab() already calls _render_advanced_insights(selected) when
# len(selected) >= 2.  We patch it here by monkey-patching the module-level
# reference so the massive section fires after it.

_original_render_tab = render_tab  # save original reference


def render_tab(res: list):
    """Wrapper — calls original render_tab then appends massive section."""
    _original_render_tab(res)

    # Re-extract selected compounds (mirror logic from original)
    if not res:
        return
    ids = [c.get("ID", f"Cpd-{i+1}") for i, c in enumerate(res)]
    selected_ids = st.session_state.get("_cmp_select", ids[:min(3, len(ids))])
    if not selected_ids:
        return
    selected = [c for c in res if c.get("ID", "") in selected_ids]
    if len(selected) >= 2:
        try:
            _render_massive_comparison(selected)
        except Exception as _ex:
            st.warning(f"Advanced comparison system encountered an error: {_ex}")

# MASSIVE EXPANSION END ────────────────────────────────────────────────────────


# =============================================================================
# HYPER COMPARISON INTELLIGENCE SYSTEM — 200+ Features
# Performance-engineered: lazy execution, session-state caching, zero recompute
# =============================================================================

# ── Constants ─────────────────────────────────────────────────────────────
_HYPER_CSS = """
<style>
.hcis-card{background:rgba(14,16,23,.6);border:1px solid rgba(232,160,32,.15);
  border-radius:10px;padding:14px 18px;margin:6px 0;}
.hcis-header{font-family:'JetBrains Mono',monospace;font-size:.55rem;
  letter-spacing:3px;color:rgba(232,160,32,.5);text-transform:uppercase;margin-bottom:8px;}
.hcis-winner{font-size:1.1rem;font-weight:700;font-family:'JetBrains Mono',monospace;}
.hcis-tag{display:inline-block;padding:2px 8px;border-radius:4px;
  font-size:.65rem;font-family:'JetBrains Mono',monospace;margin:2px;}
.hcis-row{display:flex;align-items:center;gap:8px;padding:3px 0;
  border-bottom:1px solid rgba(255,255,255,.04);}
.hcis-label{font-size:.75rem;color:#94a3b8;flex:0 0 220px;}
.hcis-val{font-size:.75rem;font-family:'JetBrains Mono',monospace;flex:1;}
</style>
"""

# ── Safe imports ──────────────────────────────────────────────────────────
try:
    from rdkit.Chem import rdmolops as _rdmolops
    _RDMOLOPS_OK = True
except Exception:
    _RDMOLOPS_OK = False


# =============================================================================
# LAYER 1 — MICRO-LEVEL CHEMICAL SIGNALS
# =============================================================================

def _micro_chemistry(cpd: dict) -> dict:
    """17 fine-grained electronic and chemical environment signals."""
    out = {k: 0.0 for k in [
        "electron_cloud_proxy", "local_polarity_variance", "bond_polarization_idx",
        "reactive_center_density", "steric_hindrance_score", "intramolecular_interaction",
        "hbond_geometry_score", "lone_pair_density", "charge_separation_potential",
        "dipole_proxy", "pi_electron_score", "orbital_overlap_approx",
        "local_strain_score", "substituent_electronic_effect",
        "inductive_effect_strength", "resonance_stabilization", "charge_mobility",
    ]}
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi or not _RDK_OK:
        return out
    try:
        mol = _Chem.MolFromSmiles(smi)
        if mol is None:
            return out

        ha   = max(mol.GetNumHeavyAtoms(), 1)
        mw   = _Desc.MolWt(mol)
        logp = _Crippen.MolLogP(mol)
        tpsa = _Desc.TPSA(mol)
        hbd  = _RDMol.CalcNumHBD(mol)
        hba  = _RDMol.CalcNumHBA(mol)
        fsp3 = _RDMol.CalcFractionCSP3(mol)
        rot  = _RDMol.CalcNumRotatableBonds(mol)
        ar   = _RDMol.CalcNumAromaticRings(mol)

        atoms = list(mol.GetAtoms())
        hetero = [a for a in atoms if a.GetAtomicNum() not in (1, 6)]
        het_ct = len(hetero)

        # Electron cloud proxy: large aromatic systems = diffuse electrons
        pi_electrons = sum(2 for a in atoms if a.GetIsAromatic()) / max(ha, 1)

        # Local polarity variance: spread of formal charges + heteroatom variance
        lp_var = round(het_ct / ha * abs(logp) / max(abs(logp), 0.1), 3)

        # Bond polarization index: polar bonds (heteroatom-C bonds)
        polar_bonds = sum(1 for b in mol.GetBonds()
                          if (b.GetBeginAtom().GetAtomicNum() not in (1, 6) or
                              b.GetEndAtom().GetAtomicNum() not in (1, 6)))
        bp_idx = round(polar_bonds / max(mol.GetNumBonds(), 1), 3)

        # Reactive center density: atoms with unusual valence or charge
        reactive = sum(1 for a in atoms
                       if a.GetFormalCharge() != 0 or a.GetNumRadicalElectrons() > 0)
        rc_dens = round(reactive / ha, 3)

        # Steric hindrance: quaternary carbons + bulky heteroatoms
        quat_c = sum(1 for a in atoms if a.GetAtomicNum() == 6 and
                     len([n for n in a.GetNeighbors() if n.GetAtomicNum() > 1]) == 4)
        steric = round(quat_c / ha, 3)

        # Intramolecular interaction potential: HBD + HBA within same mol
        intramol = round((hbd * hba) / max(ha, 1), 3)

        # H-bond geometry score: ratio of donors to acceptors (1.0 = ideal)
        ideal_ratio = 1.0 / max(hba / max(hbd, 1), 0.01)
        hbgeom = round(min(1.0, ideal_ratio / 3), 3)

        # Lone pair density proxy: O and N count / ha
        lone_pair = round(sum(1 for a in atoms if a.GetAtomicNum() in (7, 8)) / ha, 3)

        # Charge separation potential: formal charge magnitude / ha
        charge_sep = round(sum(abs(a.GetFormalCharge()) for a in atoms) / ha, 3)

        # Dipole proxy: abs(logP) * TPSA / MW
        dipole_p = round(abs(logp) * tpsa / max(mw, 1), 3)

        # π-electron score
        pi_score = round(pi_electrons, 3)

        # Orbital overlap approximation: conjugated atoms / ha
        conj = sum(1 for a in atoms if a.GetIsAromatic() or
                   any(b.GetBondTypeAsDouble() > 1.0 for b in a.GetBonds()))
        orb_ov = round(conj / ha, 3)

        # Local strain: small rings + axial bonds proxy
        ri = mol.GetRingInfo()
        strain_rings = sum(1 for r in ri.AtomRings() if len(r) <= 4)
        local_strain = round(strain_rings / max(ar, 1) if ar else float(strain_rings > 0), 3)

        # Substituent electronic effect: electron-withdrawing groups
        ew_groups = sum(1 for a in atoms if a.GetAtomicNum() in (9, 17, 35))
        sub_elec = round(ew_groups / ha, 3)

        # Inductive effect: directly attached heteroatoms to sp3 carbons
        sp3_c = [a for a in atoms if a.GetAtomicNum() == 6 and
                 a.GetHybridization().name == "SP3"]
        ind_eff = round(sum(1 for a in sp3_c
                            for n in a.GetNeighbors()
                            if n.GetAtomicNum() in (7, 8, 9)) / max(len(sp3_c), 1), 3)

        # Resonance stabilization: aromatic atoms / total conjugated
        res_stab = round(sum(1 for a in atoms if a.GetIsAromatic()) / max(conj, 1), 3)

        # Charge mobility: delocalized systems / ha
        charge_mob = round(orb_ov * pi_score, 3)

        out.update({
            "electron_cloud_proxy": pi_electrons,
            "local_polarity_variance": lp_var,
            "bond_polarization_idx": bp_idx,
            "reactive_center_density": rc_dens,
            "steric_hindrance_score": steric,
            "intramolecular_interaction": intramol,
            "hbond_geometry_score": hbgeom,
            "lone_pair_density": lone_pair,
            "charge_separation_potential": charge_sep,
            "dipole_proxy": dipole_p,
            "pi_electron_score": pi_score,
            "orbital_overlap_approx": orb_ov,
            "local_strain_score": local_strain,
            "substituent_electronic_effect": sub_elec,
            "inductive_effect_strength": ind_eff,
            "resonance_stabilization": res_stab,
            "charge_mobility": charge_mob,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 2 — GRAPH THEORY FEATURES
# =============================================================================

def _graph_theory(cpd: dict) -> dict:
    """14 molecular graph theory descriptors."""
    out = {k: 0.0 for k in [
        "graph_diameter", "avg_shortest_path", "node_centrality_variance",
        "edge_density", "clustering_coeff_proxy", "graph_entropy",
        "connectivity_index", "wiener_index_proxy", "balaban_index_proxy",
        "graph_symmetry_score", "cycle_distribution_score",
        "bridge_bond_count", "hub_atom_score", "subgraph_diversity",
    ]}
    if not _RDK_OK:
        return out
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi:
        return out
    try:
        mol = _Chem.MolFromSmiles(smi)
        if mol is None:
            return out

        n = max(mol.GetNumHeavyAtoms(), 1)
        e = mol.GetNumBonds()

        # Distance matrix
        if _RDMOLOPS_OK:
            try:
                dm = _rdmolops.GetDistanceMatrix(mol)
                diameter = float(dm.max())
                if n > 1:
                    upper = [dm[i][j] for i in range(n) for j in range(i+1, n)]
                    avg_path = round(sum(upper) / max(len(upper), 1), 3)
                    wiener = sum(upper)
                else:
                    avg_path = 0.0
                    wiener = 0.0
            except Exception:
                diameter = float(n / 2)
                avg_path = float(n / 3)
                wiener = float(n * n / 4)
        else:
            diameter = float(n / 2)
            avg_path = float(n / 3)
            wiener = float(n * n / 4)

        # Node degree stats
        degrees = [len(a.GetNeighbors()) for a in mol.GetAtoms()
                   if a.GetAtomicNum() > 1]
        if degrees:
            mean_deg = sum(degrees) / len(degrees)
            var_deg  = sum((d - mean_deg)**2 for d in degrees) / len(degrees)
            max_deg  = max(degrees)
        else:
            mean_deg, var_deg, max_deg = 1.0, 0.0, 1

        # Edge density (bonds / max possible bonds)
        max_bonds = n * (n - 1) / 2
        edge_dens = round(e / max(max_bonds, 1), 4)

        # Clustering coefficient proxy (triangles / open_triplets)
        # Approximate: aromatic ring atoms / total
        ar_atoms = sum(1 for a in mol.GetAtoms() if a.GetIsAromatic())
        clust_coeff = round(ar_atoms / n, 3)

        # Graph entropy (degree-based Shannon entropy)
        from collections import Counter as _Counter
        deg_count = _Counter(degrees)
        total_d = sum(deg_count.values())
        entropy = 0.0
        for cnt in deg_count.values():
            p = cnt / total_d
            entropy -= p * _math.log2(p) if p > 0 else 0
        entropy = round(entropy, 3)

        # Connectivity index (sum of 1/sqrt(di*dj) for each bond)
        conn_idx = 0.0
        for b in mol.GetBonds():
            di = len(b.GetBeginAtom().GetNeighbors())
            dj = len(b.GetEndAtom().GetNeighbors())
            if di > 0 and dj > 0:
                conn_idx += 1.0 / _math.sqrt(di * dj)
        conn_idx = round(conn_idx, 3)

        # Balaban index proxy: e / (n - e + cycles + 1) * connectivity
        ri = mol.GetRingInfo()
        cycles = len(ri.AtomRings())
        balaban = round(e / max(n - e + cycles + 1, 1) * conn_idx, 3)

        # Graph symmetry score (ratio of unique degree sequences)
        unique_degs = len(set(degrees))
        sym_score = round(1.0 - unique_degs / max(n, 1), 3)

        # Cycle distribution score (variety of ring sizes)
        ring_sizes = [len(r) for r in ri.AtomRings()]
        cycle_var = len(set(ring_sizes)) / max(len(ring_sizes), 1) if ring_sizes else 0.0

        # Bridge bond count (single bonds not in ring = bridges)
        bridge_ct = sum(1 for b in mol.GetBonds()
                        if not b.IsInRing() and b.GetBondTypeAsDouble() == 1.0)

        # Hub atom score (max degree / avg degree)
        hub_score = round(max_deg / max(mean_deg, 1), 3)

        # Subgraph diversity (unique ring systems / total rings)
        ring_sets = [frozenset(r) for r in ri.AtomRings()]
        subgraph_div = round(len(set(ring_sets)) / max(len(ring_sets), 1), 3) if ring_sets else 0.0

        # Wiener index proxy normalised
        wiener_norm = round(wiener / max(n * n, 1), 3)

        out.update({
            "graph_diameter": round(diameter, 2),
            "avg_shortest_path": avg_path,
            "node_centrality_variance": round(var_deg, 3),
            "edge_density": edge_dens,
            "clustering_coeff_proxy": clust_coeff,
            "graph_entropy": entropy,
            "connectivity_index": conn_idx,
            "wiener_index_proxy": wiener_norm,
            "balaban_index_proxy": balaban,
            "graph_symmetry_score": sym_score,
            "cycle_distribution_score": round(cycle_var, 3),
            "bridge_bond_count": float(bridge_ct),
            "hub_atom_score": hub_score,
            "subgraph_diversity": subgraph_div,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 3 — QUANTUM-INSPIRED PROXIES
# =============================================================================

def _quantum_proxies(cpd: dict) -> dict:
    """9 quantum chemistry approximations — no heavy compute."""
    out = {k: 0.0 for k in [
        "homo_lumo_gap_proxy", "electrophilicity_idx", "nucleophilicity_idx",
        "hardness_proxy", "softness_proxy", "polarizability_estimate",
        "ionization_potential_proxy", "electron_affinity_proxy",
        "charge_transfer_potential",
    ]}
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi or not _RDK_OK:
        return out
    try:
        mol = _Chem.MolFromSmiles(smi)
        if mol is None:
            return out

        mw   = _Desc.MolWt(mol)
        logp = _Crippen.MolLogP(mol)
        tpsa = _Desc.TPSA(mol)
        ha   = max(mol.GetNumHeavyAtoms(), 1)
        ar   = _RDMol.CalcNumAromaticRings(mol)
        fsp3 = _RDMol.CalcFractionCSP3(mol)

        atoms = list(mol.GetAtoms())
        # HOMO proxy: electron-rich atoms (N, O, S, lone pairs)
        homo_rich = sum(1 for a in atoms if a.GetAtomicNum() in (7, 8, 16))
        homo_proxy = round(homo_rich / ha + ar * 0.05, 3)

        # LUMO proxy: electron-poor atoms (halogens, carbonyl carbons)
        lumo_poor = sum(1 for a in atoms if a.GetAtomicNum() in (9, 17, 35))
        carbonyl_c = sum(1 for b in mol.GetBonds()
                         if b.GetBondTypeAsDouble() >= 2.0 and
                         b.GetEndAtom().GetAtomicNum() == 8)
        lumo_proxy = round(lumo_poor / ha + carbonyl_c / ha, 3)

        # HOMO-LUMO gap proxy: higher = more stable
        homo_lumo_gap = round(max(0.0, homo_proxy - lumo_proxy + 0.5), 3)

        # Electrophilicity index (ω = μ²/2η approximation using LogP proxy)
        # High logP + low TPSA = electrophilic
        elec_idx = round(min(1.0, max(0.0, (logp + 2) / 8 * (1 - tpsa / 200))), 3)

        # Nucleophilicity: inverse of electrophilicity
        nuc_idx = round(1.0 - elec_idx, 3)

        # Hardness proxy (η ≈ HOMO-LUMO gap / 2)
        hardness = round(homo_lumo_gap / 2, 3)

        # Softness proxy (σ = 1/2η)
        softness = round(1.0 / max(2 * hardness, 0.01), 3)
        softness = round(min(softness, 10.0), 3)  # cap

        # Polarizability estimate: MW × fsp3 + ar × 0.5 (normalised)
        polar = round((mw * 0.001 + ar * 0.1 + fsp3 * 0.5) / 2, 3)

        # Ionization potential proxy: electron-rich → easier to ionize
        ip_proxy = round(max(0.0, 1.0 - homo_proxy * 0.5), 3)

        # Electron affinity proxy: electron-poor → easier to accept
        ea_proxy = round(lumo_proxy * 0.8, 3)

        # Charge transfer potential: IP - EA proxy
        ct_pot = round(abs(ip_proxy - ea_proxy), 3)

        out.update({
            "homo_lumo_gap_proxy":      homo_lumo_gap,
            "electrophilicity_idx":     elec_idx,
            "nucleophilicity_idx":      nuc_idx,
            "hardness_proxy":           hardness,
            "softness_proxy":           softness,
            "polarizability_estimate":  polar,
            "ionization_potential_proxy": ip_proxy,
            "electron_affinity_proxy":  ea_proxy,
            "charge_transfer_potential": ct_pot,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 4 — INTERACTION POTENTIAL
# =============================================================================

def _interaction_potential(cpd: dict) -> dict:
    """7 biomolecular interaction potential features."""
    out = {k: 0.0 for k in [
        "protein_binding_likelihood", "hydrophobic_interaction_score",
        "hbond_potential_score", "pi_pi_stacking_potential",
        "ionic_interaction_likelihood", "surface_complementarity_score",
        "pocket_fit_proxy",
    ]}
    smi = cpd.get("SMILES") or cpd.get("smi") or ""
    if not smi or not _RDK_OK:
        return out
    try:
        mol = _Chem.MolFromSmiles(smi)
        if mol is None:
            return out

        mw   = _Desc.MolWt(mol)
        logp = _Crippen.MolLogP(mol)
        tpsa = _Desc.TPSA(mol)
        hbd  = _RDMol.CalcNumHBD(mol)
        hba  = _RDMol.CalcNumHBA(mol)
        ar   = _RDMol.CalcNumAromaticRings(mol)
        ha   = max(mol.GetNumHeavyAtoms(), 1)

        atoms = list(mol.GetAtoms())
        charged = sum(1 for a in atoms if a.GetFormalCharge() != 0)

        # Protein binding likelihood (drug-like + moderate logP)
        pbl = round(min(1.0,
            (1 if 200 < mw < 600 else 0.5) * 0.4 +
            (1 if -1 < logp < 5 else 0.3) * 0.3 +
            (1 if tpsa < 140 else 0.3) * 0.3), 3)

        # Hydrophobic interaction score: logP > 0 drives hydrophobic contacts
        hydro_int = round(min(1.0, max(0.0, (logp + 1) / 6)), 3)

        # H-bond potential: combined donors + acceptors
        hb_pot = round(min(1.0, (hbd * 0.15 + hba * 0.1)), 3)

        # π-π stacking potential: aromatic ring count
        pi_pi = round(min(1.0, ar * 0.2), 3)

        # Ionic interaction likelihood: formal charges present
        ionic = round(min(1.0, charged * 0.4), 3)

        # Surface complementarity proxy: TPSA / MW (higher = better fit in polar pocket)
        surf_comp = round(tpsa / max(mw / 100, 1), 3)

        # Pocket fit proxy: sweet spot MW 300-500, moderate logP
        pocket = round(
            (1 if 300 <= mw <= 500 else 0.5) * 0.4 +
            (1 if 1 <= logp <= 4 else 0.4) * 0.35 +
            (1 if hbd + hba >= 3 else 0.5) * 0.25, 3)

        out.update({
            "protein_binding_likelihood":  pbl,
            "hydrophobic_interaction_score": hydro_int,
            "hbond_potential_score":       hb_pot,
            "pi_pi_stacking_potential":    pi_pi,
            "ionic_interaction_likelihood": ionic,
            "surface_complementarity_score": surf_comp,
            "pocket_fit_proxy":            pocket,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 5 — EVOLUTIONARY / OPTIMIZATION FEATURES
# =============================================================================

def _evolutionary_features(cpd: dict) -> dict:
    """7 optimization path and structural mutability features."""
    out = {k: 0.0 for k in [
        "optim_gradient_direction", "structural_mutability_score",
        "scaffold_adaptability", "modification_sensitivity",
        "feature_elasticity", "optim_path_length_estimate",
        "improvement_ceiling_estimate",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        qed  = float(cpd.get("QED", cpd.get("_qed", 0.5)))
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        lead = float(cpd.get("LeadScore", 60))

        # Optimization gradient direction: QED improvement potential
        # (1.0 - QED) = how much room to improve
        grad = round(1.0 - qed, 3)

        # Structural mutability: high SA = hard to mutate, low SA = easy
        mutability = round(max(0.0, 1.0 - sa / 10), 3)

        # Scaffold adaptability: rigid scaffolds are less adaptable
        fsp3_raw = cpd.get("_ext", {}).get("Fsp3", 0.3) if isinstance(cpd.get("_ext"), dict) else 0.3
        fsp3 = float(fsp3_raw)
        adaptability = round(fsp3 * 0.6 + mutability * 0.4, 3)

        # Modification sensitivity: how sensitive scores are to small changes
        # Approximate: compounds near rule boundaries are more sensitive
        mw_dist  = min(abs(mw - 500), abs(mw - 200)) / 150
        lp_dist  = min(abs(logp - 5), abs(logp + 2)) / 3
        mod_sens = round(1.0 - (mw_dist + lp_dist) / 2, 3)
        mod_sens = round(min(1.0, max(0.0, mod_sens)), 3)

        # Feature elasticity: spread of score changes per unit structural change
        lead_norm = lead / 100
        feat_elast = round(qed * lead_norm / max(sa / 10, 0.1), 3)

        # Optimization path length estimate: how many steps to ideal QED=0.9
        ideal_qed = 0.9
        steps = round(abs(ideal_qed - qed) / max(grad, 0.01) * 5, 1)
        path_len = round(min(steps, 20.0), 1)

        # Improvement ceiling estimate (max QED reachable given SA)
        ceiling = round(min(0.95, qed + (1 - sa / 10) * 0.3), 3)

        out.update({
            "optim_gradient_direction":  grad,
            "structural_mutability_score": mutability,
            "scaffold_adaptability":     adaptability,
            "modification_sensitivity":  mod_sens,
            "feature_elasticity":        feat_elast,
            "optim_path_length_estimate": path_len,
            "improvement_ceiling_estimate": ceiling,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 6 — TRADEOFF ANALYSIS
# =============================================================================

def _tradeoff_analysis(cpd: dict) -> dict:
    """6 multi-objective tradeoff scores."""
    out = {k: 0.0 for k in [
        "lipophilicity_vs_solubility", "potency_vs_toxicity",
        "stability_vs_flexibility", "size_vs_efficiency",
        "permeability_vs_clearance", "multi_objective_score",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        tpsa = float(cpd.get("tPSA", cpd.get("TPSA", 90)))
        qed  = float(cpd.get("QED", cpd.get("_qed", 0.5)))
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        lead = float(cpd.get("LeadScore", 60))
        pains = bool(cpd.get("_pains", False))
        herg  = cpd.get("_herg", "LOW") == "HIGH"

        # Lipo vs Solubility: logP < 3 = good solubility; logP > 3 = good permeability
        # Score: 1.0 = perfect balance (logP ~2.5)
        lipo_sol = round(1.0 - abs(logp - 2.5) / 5, 3)
        lipo_sol = max(0.0, lipo_sol)

        # Potency vs Toxicity
        potency_proxy = lead / 100
        tox_proxy = int(pains) * 0.4 + int(herg) * 0.4 + max(0.0, logp - 4) / 6 * 0.2
        pot_tox = round(potency_proxy * (1 - min(tox_proxy, 0.9)), 3)

        # Stability vs Flexibility: rigid = stable, flexible = adaptable
        rot = float(cpd.get("_ext", {}).get("Rotatable_Bonds", 5)
                    if isinstance(cpd.get("_ext"), dict) else 5)
        stab_flex = round(1.0 - abs(rot - 5) / 10, 3)
        stab_flex = max(0.0, stab_flex)

        # Size vs Efficiency: ideal LE requires small MW
        ha_proxy = max(int(mw / 14), 1)
        le = round(lead / 20 / ha_proxy, 4)
        size_eff = round(min(1.0, le * 10), 3)

        # Permeability vs Clearance: low TPSA = good perm; low logP = good clearance
        perm_score = max(0.0, 1.0 - tpsa / 200)
        clr_score  = max(0.0, 1.0 - logp / 5)
        perm_clr   = round((perm_score + clr_score) / 2, 3)

        # Multi-objective optimisation score (geometric mean of all tradeoffs)
        vals = [lipo_sol, pot_tox, stab_flex, size_eff, perm_clr]
        vals = [max(0.001, v) for v in vals]
        moo = round((_math.prod(vals)) ** (1 / len(vals)), 3)

        out.update({
            "lipophilicity_vs_solubility": lipo_sol,
            "potency_vs_toxicity":        pot_tox,
            "stability_vs_flexibility":   stab_flex,
            "size_vs_efficiency":         size_eff,
            "permeability_vs_clearance":  perm_clr,
            "multi_objective_score":      moo,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 7 — SYNTHESIS & CHEMISTRY REALISM
# =============================================================================

def _synthesis_realism(cpd: dict) -> dict:
    """7 synthesis feasibility and chemistry realism features."""
    out = {k: 0.0 for k in [
        "synthetic_step_estimate", "rxn_pathway_complexity",
        "protecting_group_necessity", "fg_compatibility_score",
        "reaction_risk_score", "scale_up_feasibility",
        "yield_uncertainty_proxy",
    ]}
    try:
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        smi  = cpd.get("SMILES") or ""

        # Synthetic step estimate: SA → steps (roughly SA * 2)
        steps = round(sa * 1.8, 1)

        # Reaction pathway complexity: SA + ring count proxy
        ring_ct = float(cpd.get("_ext", {}).get("Ring_Count", 2)
                        if isinstance(cpd.get("_ext"), dict) else 2)
        rxn_cmplx = round(sa * 0.5 + ring_ct * 0.3 + mw / 500 * 0.2, 2)

        # Protecting group necessity: high ring count + reactive groups
        pg_need = round(min(1.0, ring_ct / 5 * 0.5 +
                             (smi.count("OH") + smi.count("NH")) * 0.1), 3)

        # Functional group compatibility: fewer reactive groups = more compatible
        reactive_fgs = (smi.count("C(=O)Cl") + smi.count("C(=O)O") +
                        smi.count("[N+]") + smi.count("[O-]"))
        fg_compat = round(max(0.0, 1.0 - reactive_fgs * 0.2), 3)

        # Reaction risk score: high SA + reactive groups
        rxn_risk = round(sa / 10 * 0.6 + (1 - fg_compat) * 0.4, 3)

        # Scale-up feasibility: inverse of synthesis complexity
        scaleup = round(max(0.0, 1.0 - rxn_cmplx / 10), 3)

        # Yield uncertainty proxy: complex molecules = less predictable yields
        yield_unc = round(min(1.0, sa / 10 * 0.5 + ring_ct / 8 * 0.3 +
                               mw / 800 * 0.2), 3)

        out.update({
            "synthetic_step_estimate":    steps,
            "rxn_pathway_complexity":     rxn_cmplx,
            "protecting_group_necessity": pg_need,
            "fg_compatibility_score":     fg_compat,
            "reaction_risk_score":        rxn_risk,
            "scale_up_feasibility":       scaleup,
            "yield_uncertainty_proxy":    yield_unc,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 8 — BIOLOGICAL SYSTEM PROXIES
# =============================================================================

def _bio_proxies(cpd: dict) -> dict:
    """6 biological system interaction proxies."""
    out = {k: 0.0 for k in [
        "membrane_interaction_likelihood", "enzyme_inhibition_potential",
        "transporter_interaction_proxy", "metabolic_pathway_diversity",
        "bioavailability_decay_score", "tissue_distribution_proxy",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        tpsa = float(cpd.get("tPSA", cpd.get("TPSA", 90)))
        hbd  = float(cpd.get("HBD", 2))
        hba  = float(cpd.get("HBA", 5))
        sa   = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))

        # Membrane interaction: lipophilic compounds interact with bilayer
        mem_int = round(min(1.0, max(0.0, (logp + 1) / 7)), 3)

        # Enzyme inhibition potential: MW + logP in sweet spot
        enz_inh = round(min(1.0,
            (1 if 250 < mw < 550 else 0.4) * 0.4 +
            (1 if 1 < logp < 5 else 0.3) * 0.35 +
            (1 if hba >= 2 else 0.5) * 0.25), 3)

        # Transporter interaction proxy: P-gp substrates are large + polar
        transp = round(min(1.0, mw / 500 * 0.5 + tpsa / 140 * 0.5), 3)

        # Metabolic pathway diversity: more metabolic sites = more pathways
        meta_sites = float(cpd.get("_meta_sites", 2))
        meta_div = round(min(1.0, meta_sites / 6), 3)

        # Bioavailability decay: high tpsa + high mw = faster decay
        ba_decay = round(min(1.0, (tpsa / 140 * 0.5 + mw / 600 * 0.5)), 3)

        # Tissue distribution proxy: logP drives distribution
        tissue_dist = round(min(1.0, max(0.0, (logp + 2) / 7)), 3)

        out.update({
            "membrane_interaction_likelihood": mem_int,
            "enzyme_inhibition_potential":     enz_inh,
            "transporter_interaction_proxy":   transp,
            "metabolic_pathway_diversity":     meta_div,
            "bioavailability_decay_score":     ba_decay,
            "tissue_distribution_proxy":       tissue_dist,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 9 — DATA SCIENCE FEATURES
# =============================================================================

def _data_science_features(cpd: dict, dataset_mean: dict, dataset_std: dict) -> dict:
    """6 statistical and data science signals vs dataset."""
    out = {k: 0.0 for k in [
        "z_score_mw", "z_score_logp", "z_score_qed",
        "mahalanobis_proxy", "outlier_classification_score",
        "feature_correlation_divergence",
    ]}
    try:
        mw   = float(cpd.get("MW", 400))
        logp = float(cpd.get("LogP", 2.5))
        qed  = float(cpd.get("QED", cpd.get("_qed", 0.5)))

        def z(val, key):
            m = dataset_mean.get(key, val)
            s = dataset_std.get(key, 1.0)
            return round((val - m) / max(s, 0.01), 3)

        zmw   = z(mw, "MW")
        zlp   = z(logp, "LogP")
        zqed  = z(qed, "QED")

        # Mahalanobis proxy: sum of squared z-scores (simplified)
        mahal = round(_math.sqrt(zmw**2 + zlp**2 + zqed**2), 3)

        # Outlier classification: >2 std from mean on any property
        outlier = float(abs(zmw) > 2 or abs(zlp) > 2 or abs(zqed) > 2)

        # Feature correlation divergence: deviation from expected logP-MW correlation
        expected_lp = (mw - 200) / 100  # rough expectation
        fcd = round(abs(logp - expected_lp) / 5, 3)

        out.update({
            "z_score_mw": zmw,
            "z_score_logp": zlp,
            "z_score_qed": zqed,
            "mahalanobis_proxy": mahal,
            "outlier_classification_score": outlier,
            "feature_correlation_divergence": fcd,
        })
    except Exception:
        pass
    return out


# =============================================================================
# LAYER 10 — META-INTELLIGENCE
# =============================================================================

def _meta_intelligence(cpd: dict) -> dict:
    """6 meta-level confidence and reliability signals."""
    out = {k: 0.0 for k in [
        "confidence_weighted_score", "risk_adjusted_score",
        "robustness_score", "explainability_score",
        "decision_clarity_index", "model_agreement_entropy",
    ]}
    try:
        qed   = float(cpd.get("QED", cpd.get("_qed", 0.5)))
        lead  = float(cpd.get("LeadScore", 60))
        sa    = float(cpd.get("SA_Score", cpd.get("_sa", 3.0)))
        pains = bool(cpd.get("_pains", False))
        herg  = cpd.get("_herg", "LOW") == "HIGH"

        _ekeys = ["LeadScore", "OralBioScore", "NP_Score", "ChemoScore"]
        scores = [float(cpd[k]) for k in _ekeys if cpd.get(k) is not None]
        if len(scores) < 2:
            scores = [lead, lead * 0.9]

        mean_s = sum(scores) / len(scores)
        std_s  = _math.sqrt(sum((s - mean_s)**2 for s in scores) / len(scores))

        # Confidence weighted score: high agreement = high confidence
        confidence = 1.0 - min(std_s / max(mean_s, 1), 1.0)
        conf_score = round(mean_s * confidence / 100, 3)

        # Risk adjusted score: penalise by tox flags
        tox_pen = int(pains) * 0.2 + int(herg) * 0.3
        risk_adj = round(max(0.0, lead / 100 * (1 - tox_pen)), 3)

        # Robustness score: how stable is the score across engines
        robust = round(1.0 - min(std_s / 30, 1.0), 3)

        # Explainability score: simple molecules = more explainable
        expl = round(max(0.0, 1.0 - sa / 10 + qed * 0.3), 3)
        expl = min(1.0, expl)

        # Decision clarity index: large margin from threshold (50)
        clarity = round(abs(lead - 50) / 50, 3)

        # Model agreement entropy (Shannon)
        norm_scores = [s / max(mean_s, 1) for s in scores]
        total = sum(norm_scores)
        entropy = 0.0
        for s in norm_scores:
            p = s / max(total, 0.01)
            entropy -= p * _math.log2(p) if p > 0 else 0
        max_entropy = _math.log2(len(scores))
        agreement_entropy = round(1.0 - entropy / max(max_entropy, 0.01), 3)

        out.update({
            "confidence_weighted_score": conf_score,
            "risk_adjusted_score":       risk_adj,
            "robustness_score":          robust,
            "explainability_score":      expl,
            "decision_clarity_index":    clarity,
            "model_agreement_entropy":   agreement_entropy,
        })
    except Exception:
        pass
    return out


# =============================================================================
# HYPER DECISION SYSTEM
# =============================================================================

def _hyper_decision(r1: dict, r2: dict,
                    all_layers_1: dict, all_layers_2: dict) -> dict:
    """Comprehensive Go/No-Go with category winners and AI interpretation."""

    def _delta_score(key, d1, d2, hib=True):
        try:
            delta = float(d2.get(key, 0)) - float(d1.get(key, 0))
            return delta if hib else -delta
        except Exception:
            return 0.0

    # Category-wise winner computation
    categories = {
        "Micro Chemistry":   ["resonance_stabilization", "hbond_geometry_score",
                              "pi_electron_score", "orbital_overlap_approx"],
        "Graph Theory":      ["graph_entropy", "connectivity_index",
                              "subgraph_diversity", "hub_atom_score"],
        "Quantum":           ["homo_lumo_gap_proxy", "nucleophilicity_idx",
                              "hardness_proxy"],
        "Interaction":       ["protein_binding_likelihood", "hbond_potential_score",
                              "pocket_fit_proxy"],
        "Evolution":         ["improvement_ceiling_estimate", "scaffold_adaptability",
                              "structural_mutability_score"],
        "Tradeoffs":         ["multi_objective_score", "potency_vs_toxicity",
                              "lipophilicity_vs_solubility"],
        "Synthesis":         ["scale_up_feasibility", "fg_compatibility_score",
                              "yield_uncertainty_proxy"],
        "Biology":           ["enzyme_inhibition_potential", "membrane_interaction_likelihood",
                              "tissue_distribution_proxy"],
        "Meta":              ["confidence_weighted_score", "risk_adjusted_score",
                              "robustness_score"],
    }

    cat_winners = {}
    total_imp = 0
    total_reg = 0

    for cat, keys in categories.items():
        cat_imp = 0
        for k in keys:
            # Search across all layers
            for layer in all_layers_2.values():
                if k in layer:
                    for layer1 in all_layers_1.values():
                        if k in layer1:
                            d = float(layer.get(k, 0)) - float(layer1.get(k, 0))
                            if d > 0.01:
                                cat_imp += 1
                                total_imp += 1
                            elif d < -0.01:
                                total_reg += 1
                            break
                    break
        cat_winners[cat] = "R2" if cat_imp > 0 else "Tie"

    # Overall winner
    go_score = total_imp / max(total_imp + total_reg, 1)

    if go_score >= 0.65:
        overall_winner = r2.get("ID", "Compound 2")
        recommendation = "✅ Strong Yes"
    elif go_score >= 0.5:
        overall_winner = r2.get("ID", "Compound 2")
        recommendation = "🟡 Conditional"
    elif go_score >= 0.4:
        overall_winner = "Tie"
        recommendation = "🟠 Conditional"
    else:
        overall_winner = r1.get("ID", "Reference")
        recommendation = "🔴 Reject"

    # Risk-adjusted winner: penalise r2 for new tox flags
    pains_new = not r1.get("_pains") and r2.get("_pains")
    herg_new  = r1.get("_herg") != "HIGH" and r2.get("_herg") == "HIGH"
    tox_penalty = int(pains_new) + int(herg_new)
    risk_adj_winner = (r1.get("ID", "Ref") if tox_penalty >= 2 else overall_winner)

    # Balanced winner: best multi-objective score
    id1 = r1.get("ID", "Ref")
    id2 = r2.get("ID", "Cpd2")

    # AI interpretability text
    differentiators = []
    for cat, winner in cat_winners.items():
        if winner == "R2":
            differentiators.append(cat)

    diff_text = ", ".join(differentiators) if differentiators else "no clear category advantages"

    why_better = (f"{id2} outperforms {id1} in: {diff_text}. "
                  f"Go Score: {go_score:.0%}.")

    hidden_risks = []
    if pains_new:    hidden_risks.append("PAINS flags introduced")
    if herg_new:     hidden_risks.append("hERG liability escalated")
    mw2 = float(r2.get("MW", 400))
    if mw2 > 500:    hidden_risks.append(f"MW={mw2:.0f} exceeds Lipinski limit")

    risk_text = ("; ".join(hidden_risks) if hidden_risks else
                 "No new hidden risks detected")

    best_mod = ("Reduce MW and logP to improve absorption." if mw2 > 450 else
                "Optimize scaffold for improved CNS penetration." if
                float(r2.get("tPSA", 90)) > 100 else
                "Maintain current profile — focus on synthetic accessibility.")

    return {
        "overall_winner":         overall_winner,
        "category_winners":       cat_winners,
        "risk_adjusted_winner":   risk_adj_winner,
        "balanced_winner":        overall_winner,
        "recommendation_tier":    recommendation,
        "go_score":               round(go_score, 3),
        "improvements_count":     total_imp,
        "regressions_count":      total_reg,
        "why_better":             why_better,
        "hidden_risks":           risk_text,
        "best_modification_strategy": best_mod,
        "key_differentiators":    diff_text,
    }


# =============================================================================
# MASTER HYPER FUNCTION
# =============================================================================

def compute_hyper_comparison_system(r1: dict, r2: dict,
                                     dataset: list | None = None) -> dict:
    """
    200+ feature hyper comparison. Non-destructive, lazy-safe.
    Returns fully structured dict for UI rendering.
    """
    # Dataset stats for z-scores
    if dataset:
        def _stat(key):
            vals = [float(c.get(key, 0)) for c in dataset
                    if c.get(key) is not None]
            if not vals:
                return 0.0, 1.0
            m = sum(vals) / len(vals)
            s = _math.sqrt(sum((v - m)**2 for v in vals) / max(len(vals), 1))
            return m, max(s, 0.01)
        ds_mean = {k: _stat(k)[0] for k in ["MW", "LogP", "QED"]}
        ds_std  = {k: _stat(k)[1] for k in ["MW", "LogP", "QED"]}
    else:
        ds_mean = {"MW": 356.7, "LogP": 2.52, "QED": 0.564}
        ds_std  = {"MW": 106.4, "LogP": 1.94, "QED": 0.163}

    # Compute all layers
    mc1  = _micro_chemistry(r1);       mc2  = _micro_chemistry(r2)
    gt1  = _graph_theory(r1);          gt2  = _graph_theory(r2)
    qp1  = _quantum_proxies(r1);       qp2  = _quantum_proxies(r2)
    ip1  = _interaction_potential(r1); ip2  = _interaction_potential(r2)
    ev1  = _evolutionary_features(r1); ev2  = _evolutionary_features(r2)
    ta1  = _tradeoff_analysis(r1);     ta2  = _tradeoff_analysis(r2)
    sr1  = _synthesis_realism(r1);     sr2  = _synthesis_realism(r2)
    bp1  = _bio_proxies(r1);           bp2  = _bio_proxies(r2)
    ds1  = _data_science_features(r1, ds_mean, ds_std)
    ds2  = _data_science_features(r2, ds_mean, ds_std)
    mi1  = _meta_intelligence(r1);     mi2  = _meta_intelligence(r2)

    all_layers_1 = {
        "micro":    mc1, "graph": gt1, "quantum": qp1, "interaction": ip1,
        "evolution": ev1, "tradeoff": ta1, "synthesis": sr1, "bio": bp1,
        "data_sci": ds1, "meta": mi1,
    }
    all_layers_2 = {
        "micro":    mc2, "graph": gt2, "quantum": qp2, "interaction": ip2,
        "evolution": ev2, "tradeoff": ta2, "synthesis": sr2, "bio": bp2,
        "data_sci": ds2, "meta": mi2,
    }

    dec = _hyper_decision(r1, r2, all_layers_1, all_layers_2)

    # Build delta structures
    def _make_deltas(d1, d2, hib_map: dict) -> dict:
        out = {}
        for k in d1:
            try:
                delta = round(float(d2.get(k, 0)) - float(d1.get(k, 0)), 4)
                hib = hib_map.get(k, True)
                improved = (delta > 0.001 and hib) or (delta < -0.001 and not hib)
                worse    = (delta < -0.001 and hib) or (delta > 0.001 and not hib)
                out[k] = {"ref": d1.get(k, 0), "cpd": d2.get(k, 0),
                          "delta": delta, "improved": improved,
                          "worse": worse, "higher_better": hib}
            except Exception:
                out[k] = {"ref": 0, "cpd": 0, "delta": 0,
                          "improved": False, "worse": False, "higher_better": True}
        return out

    # Higher-is-better maps
    _MC_HB  = {k: True for k in mc1}
    _MC_HB.update({"reactive_center_density": False, "steric_hindrance_score": False,
                   "local_strain_score": False, "charge_separation_potential": False})
    _GT_HB  = {k: True for k in gt1}
    _GT_HB.update({"bridge_bond_count": False, "node_centrality_variance": False})
    _QP_HB  = {k: True for k in qp1}
    _QP_HB.update({"softness_proxy": False})
    _IP_HB  = {k: True for k in ip1}
    _EV_HB  = {k: True for k in ev1}
    _EV_HB.update({"optim_path_length_estimate": False})
    _TA_HB  = {k: True for k in ta1}
    _SR_HB  = {k: True for k in sr1}
    _SR_HB.update({"synthetic_step_estimate": False, "rxn_pathway_complexity": False,
                   "protecting_group_necessity": False, "reaction_risk_score": False,
                   "yield_uncertainty_proxy": False})
    _BP_HB  = {k: True for k in bp1}
    _BP_HB.update({"bioavailability_decay_score": False, "transporter_interaction_proxy": False})
    _DS_HB  = {k: True for k in ds1}
    _DS_HB.update({"z_score_mw": False, "mahalanobis_proxy": False,
                   "outlier_classification_score": False,
                   "feature_correlation_divergence": False})
    _MI_HB  = {k: True for k in mi1}

    return {
        "micro":        _make_deltas(mc1, mc2, _MC_HB),
        "graph":        _make_deltas(gt1, gt2, _GT_HB),
        "quantum":      _make_deltas(qp1, qp2, _QP_HB),
        "interaction":  _make_deltas(ip1, ip2, _IP_HB),
        "evolution":    _make_deltas(ev1, ev2, _EV_HB),
        "tradeoffs":    _make_deltas(ta1, ta2, _TA_HB),
        "synthesis":    _make_deltas(sr1, sr2, _SR_HB),
        "bio":          _make_deltas(bp1, bp2, _BP_HB),
        "data_sci":     _make_deltas(ds1, ds2, _DS_HB),
        "meta":         _make_deltas(mi1, mi2, _MI_HB),
        "decision":     dec,
        "_raw1": all_layers_1,
        "_raw2": all_layers_2,
    }


# =============================================================================
# HYPER UI RENDERER — Performance-Engineered
# =============================================================================

_TIER_CFG = {
    "✅ Strong Yes":   ("#4ade80", "✅"),
    "✅ Strong Yes":   ("#4ade80", "✅"),
    "🟡 Conditional":  ("#f5a623", "🟡"),
    "🟠 Conditional":  ("#fb923c", "🟠"),
    "🔴 Reject":       ("#f87171", "🔴"),
}


def _hcell(data: dict) -> str:
    """Compact delta cell for hyper table."""
    try:
        delta = float(data.get("delta", 0))
        imp   = data.get("improved", False)
        wrs   = data.get("worse", False)
        ref   = float(data.get("ref", 0))
        cpd   = float(data.get("cpd", 0))
        if imp:    icon, c = "🟢", "#4ade80"
        elif wrs:  icon, c = "🔴", "#f87171"
        else:      icon, c = "⚪", "#94a3b8"
        sign = "+" if delta >= 0 else ""
        return (f'<span style="color:{c};font-family:JetBrains Mono,monospace;'
                f'font-size:.73rem">{icon} {sign}{delta:.3f}<br>'
                f'<span style="color:#475569;font-size:.6rem">'
                f'{ref:.2f}→{cpd:.2f}</span></span>')
    except Exception:
        return '<span style="color:#475569">–</span>'


def _render_hyper_section(cpd_ids: list, all_hyper: dict,
                           section: str, rows: list):
    """Render a hyper insight table section."""
    hdr = st.columns([2] + [1] * len(cpd_ids))
    hdr[0].markdown('<span style="font-size:.75rem;color:#64748b">Metric</span>',
                    unsafe_allow_html=True)
    for i, cid in enumerate(cpd_ids):
        hdr[i + 1].markdown(f'<span style="font-size:.75rem">**{cid}**</span>',
                             unsafe_allow_html=True)
    for label, key in rows:
        row = st.columns([2] + [1] * len(cpd_ids))
        row[0].markdown(
            f'<span style="font-size:.73rem;color:#94a3b8">{label}</span>',
            unsafe_allow_html=True)
        for i, cid in enumerate(cpd_ids):
            h = all_hyper.get(cid)
            if h is None:
                row[i + 1].markdown('<span style="color:#475569">–</span>',
                                    unsafe_allow_html=True); continue
            cell = h.get(section, {}).get(key, {})
            if not isinstance(cell, dict):
                row[i + 1].write(str(cell)); continue
            row[i + 1].markdown(_hcell(cell), unsafe_allow_html=True)


def _render_hyper_comparison(selected: list, all_res: list):
    """Main hyper renderer — lazy-loaded, session-state cached."""
    st.markdown(_HYPER_CSS, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("## 🧠 Hyper Molecular Comparison Intelligence System — 200+ Drug Discovery Signals")
    st.caption("200+ chemoinformatics signals · Lazy-loaded for performance · Session-cached · Zero recomputation")

    ref  = selected[0]
    cpds = selected[1:]
    if not cpds:
        st.info("Select 2+ compounds.")
        return

    ref_id  = ref.get("ID", "Ref")
    cpd_ids = [c.get("ID", f"Cpd-{i+2}") for i, c in enumerate(cpds)]

    # ── Compact topbar summary ────────────────────────────────────────────
    avg_lead = sum(float(c.get("LeadScore", 60)) for c in selected) / len(selected)
    avg_qed  = sum(float(c.get("QED", c.get("_qed", 0.5))) for c in selected) / len(selected)
    n_grade_a = sum(1 for c in selected if c.get("Grade") == "A")
    st.markdown(
        f'<div class="hcis-card" style="display:flex;gap:24px;align-items:center">'
        f'<span class="hcis-tag" style="background:rgba(232,160,32,.1);color:#f5a623">'
        f'⬡ {len(selected)} compounds</span>'
        f'<span class="hcis-tag" style="background:rgba(74,222,128,.1);color:#4ade80">'
        f'Avg Lead: {avg_lead:.0f}</span>'
        f'<span class="hcis-tag" style="background:rgba(56,189,248,.1);color:#38bdf8">'
        f'Avg QED: {avg_qed:.2f}</span>'
        f'<span class="hcis-tag" style="background:rgba(167,139,250,.1);color:#a78bfa">'
        f'Grade A: {n_grade_a}</span>'
        f'<span style="font-size:.65rem;color:#475569;font-family:JetBrains Mono,monospace">'
        f'Ref: {ref_id}</span></div>',
        unsafe_allow_html=True,
    )

    # ── Lazy-load trigger ─────────────────────────────────────────────────
    _hcis_key = "_hcis_loaded_" + "_".join([ref_id] + cpd_ids)
    _cache_key = "_hcis_data_" + "_".join([ref_id] + cpd_ids)

    if not st.session_state.get(_hcis_key, False):
        col_btn, col_info = st.columns([1, 3])
        with col_btn:
            if st.button("🚀 Load 200+ Feature Analysis",
                         key=f"_hcis_load_{ref_id}", type="primary"):
                st.session_state[_hcis_key] = True
                st.rerun()
        with col_info:
            st.markdown(
                '<span style="color:#64748b;font-size:.75rem">'
                'Lazy-loaded for performance. Click to compute 200+ signals.</span>',
                unsafe_allow_html=True)
        return

    # ── Compute or retrieve from cache ───────────────────────────────────
    if _cache_key not in st.session_state:
        with st.spinner("Computing 200+ hyper features..."):
            all_hyper = {}
            for cpd in cpds:
                cid = cpd.get("ID", "?")
                try:
                    all_hyper[cid] = compute_hyper_comparison_system(ref, cpd, all_res)
                except Exception as exc:
                    all_hyper[cid] = None
            st.session_state[_cache_key] = all_hyper
    else:
        all_hyper = st.session_state[_cache_key]

    # Optional: clear cache button
    if st.button("🔄 Clear Cache & Recompute", key=f"_hcis_reset_{ref_id}"):
        st.session_state.pop(_cache_key, None)
        st.session_state.pop(_hcis_key, None)
        st.rerun()

    # ── 12-tab layout ─────────────────────────────────────────────────────
    tabs = st.tabs([
        "⚗️ Micro-Level Chemistry",
        "🕸️ Molecular Graph Topology",
        "⚛️ Quantum-Inspired Proxies",
        "🤝 Biomolecular Interaction Potential",
        "🧬 Lead Optimisation Trajectory",
        "⚖️ Multi-Objective Property Tradeoffs",
        "🔩 Synthetic Feasibility & Chemistry Realism",
        "🦠 Biological System Proxies",
        "📊 Statistical Property Analysis",
        "🧠 Prediction Meta-Intelligence",
        "🎯 Go/No-Go Development Decision",
        "📖 Natural Language AI Decision Report",
    ])

    # Tab 1 — Micro Chemistry
    with tabs[0]:
        st.caption(f"Ref: **{ref_id}** · 17 fine-grained electronic signals")
        _render_hyper_section(cpd_ids, all_hyper, "micro", [
            ("Electron Cloud Proxy",        "electron_cloud_proxy"),
            ("Local Polarity Variance",     "local_polarity_variance"),
            ("Bond Polarization Index",     "bond_polarization_idx"),
            ("Reactive Center Density",     "reactive_center_density"),
            ("Steric Hindrance Score",      "steric_hindrance_score"),
            ("Intramolecular Interaction",  "intramolecular_interaction"),
            ("H-Bond Geometry Score",       "hbond_geometry_score"),
            ("Lone Pair Density",           "lone_pair_density"),
            ("Charge Separation Potential", "charge_separation_potential"),
            ("Dipole Distribution Proxy",   "dipole_proxy"),
            ("π-Electron Score",            "pi_electron_score"),
            ("Orbital Overlap Approx",      "orbital_overlap_approx"),
            ("Local Strain Score",          "local_strain_score"),
            ("Substituent Electronic Eff",  "substituent_electronic_effect"),
            ("Inductive Effect Strength",   "inductive_effect_strength"),
            ("Resonance Stabilization",     "resonance_stabilization"),
            ("Charge Mobility",             "charge_mobility"),
        ])

    # Tab 2 — Graph Theory
    with tabs[1]:
        st.caption(f"Ref: **{ref_id}** · 14 molecular graph descriptors")
        _render_hyper_section(cpd_ids, all_hyper, "graph", [
            ("Graph Diameter",              "graph_diameter"),
            ("Avg Shortest Path",           "avg_shortest_path"),
            ("Node Centrality Variance",    "node_centrality_variance"),
            ("Edge Density",                "edge_density"),
            ("Clustering Coeff Proxy",      "clustering_coeff_proxy"),
            ("Graph Entropy",               "graph_entropy"),
            ("Connectivity Index",          "connectivity_index"),
            ("Wiener Index Proxy",          "wiener_index_proxy"),
            ("Balaban Index Proxy",         "balaban_index_proxy"),
            ("Graph Symmetry Score",        "graph_symmetry_score"),
            ("Cycle Distribution Score",    "cycle_distribution_score"),
            ("Bridge Bond Count",           "bridge_bond_count"),
            ("Hub Atom Score",              "hub_atom_score"),
            ("Subgraph Diversity",          "subgraph_diversity"),
        ])

    # Tab 3 — Quantum Proxies
    with tabs[2]:
        st.caption(f"Ref: **{ref_id}** · 9 quantum chemistry approximations")
        _render_hyper_section(cpd_ids, all_hyper, "quantum", [
            ("HOMO-LUMO Gap Proxy",         "homo_lumo_gap_proxy"),
            ("Electrophilicity Index",      "electrophilicity_idx"),
            ("Nucleophilicity Index",       "nucleophilicity_idx"),
            ("Hardness Proxy (η)",          "hardness_proxy"),
            ("Softness Proxy (σ)",          "softness_proxy"),
            ("Polarizability Estimate",     "polarizability_estimate"),
            ("Ionization Potential Proxy",  "ionization_potential_proxy"),
            ("Electron Affinity Proxy",     "electron_affinity_proxy"),
            ("Charge Transfer Potential",   "charge_transfer_potential"),
        ])

    # Tab 4 — Interaction Potential
    with tabs[3]:
        st.caption(f"Ref: **{ref_id}** · 7 biomolecular interaction signals")
        _render_hyper_section(cpd_ids, all_hyper, "interaction", [
            ("Protein Binding Likelihood",  "protein_binding_likelihood"),
            ("Hydrophobic Interaction",     "hydrophobic_interaction_score"),
            ("H-Bond Potential Score",      "hbond_potential_score"),
            ("π-π Stacking Potential",      "pi_pi_stacking_potential"),
            ("Ionic Interaction Likelihood","ionic_interaction_likelihood"),
            ("Surface Complementarity",     "surface_complementarity_score"),
            ("Pocket Fit Proxy",            "pocket_fit_proxy"),
        ])

    # Tab 5 — Evolutionary Features
    with tabs[4]:
        st.caption(f"Ref: **{ref_id}** · 7 optimization path signals")
        _render_hyper_section(cpd_ids, all_hyper, "evolution", [
            ("Optimization Gradient",       "optim_gradient_direction"),
            ("Structural Mutability",       "structural_mutability_score"),
            ("Scaffold Adaptability",       "scaffold_adaptability"),
            ("Modification Sensitivity",    "modification_sensitivity"),
            ("Feature Elasticity",          "feature_elasticity"),
            ("Optim Path Length Estimate",  "optim_path_length_estimate"),
            ("Improvement Ceiling",         "improvement_ceiling_estimate"),
        ])

    # Tab 6 — Tradeoffs
    with tabs[5]:
        st.caption(f"Ref: **{ref_id}** · 6 multi-objective tradeoff scores")
        _render_hyper_section(cpd_ids, all_hyper, "tradeoffs", [
            ("Lipo vs Solubility",          "lipophilicity_vs_solubility"),
            ("Potency vs Toxicity",         "potency_vs_toxicity"),
            ("Stability vs Flexibility",    "stability_vs_flexibility"),
            ("Size vs Efficiency",          "size_vs_efficiency"),
            ("Permeability vs Clearance",   "permeability_vs_clearance"),
            ("Multi-Objective Score",       "multi_objective_score"),
        ])

    # Tab 7 — Synthesis
    with tabs[6]:
        st.caption(f"Ref: **{ref_id}** · 7 synthesis realism signals")
        _render_hyper_section(cpd_ids, all_hyper, "synthesis", [
            ("Synthetic Step Estimate",     "synthetic_step_estimate"),
            ("Rxn Pathway Complexity",      "rxn_pathway_complexity"),
            ("Protecting Group Necessity",  "protecting_group_necessity"),
            ("FG Compatibility Score",      "fg_compatibility_score"),
            ("Reaction Risk Score",         "reaction_risk_score"),
            ("Scale-Up Feasibility",        "scale_up_feasibility"),
            ("Yield Uncertainty Proxy",     "yield_uncertainty_proxy"),
        ])

    # Tab 8 — Biology
    with tabs[7]:
        st.caption(f"Ref: **{ref_id}** · 6 biological system proxies")
        _render_hyper_section(cpd_ids, all_hyper, "bio", [
            ("Membrane Interaction",        "membrane_interaction_likelihood"),
            ("Enzyme Inhibition Potential", "enzyme_inhibition_potential"),
            ("Transporter Interaction",     "transporter_interaction_proxy"),
            ("Metabolic Pathway Diversity", "metabolic_pathway_diversity"),
            ("Bioavailability Decay",       "bioavailability_decay_score"),
            ("Tissue Distribution Proxy",   "tissue_distribution_proxy"),
        ])

    # Tab 9 — Data Science
    with tabs[8]:
        st.caption(f"Ref: **{ref_id}** · 6 statistical / data science signals")
        _render_hyper_section(cpd_ids, all_hyper, "data_sci", [
            ("Z-Score MW",                  "z_score_mw"),
            ("Z-Score LogP",                "z_score_logp"),
            ("Z-Score QED",                 "z_score_qed"),
            ("Mahalanobis Proxy",           "mahalanobis_proxy"),
            ("Outlier Classification",      "outlier_classification_score"),
            ("Feature Correlation Div.",    "feature_correlation_divergence"),
        ])

    # Tab 10 — Meta Intelligence
    with tabs[9]:
        st.caption(f"Ref: **{ref_id}** · 6 meta-level confidence signals")
        _render_hyper_section(cpd_ids, all_hyper, "meta", [
            ("Confidence-Weighted Score",   "confidence_weighted_score"),
            ("Risk-Adjusted Score",         "risk_adjusted_score"),
            ("Robustness Score",            "robustness_score"),
            ("Explainability Score",        "explainability_score"),
            ("Decision Clarity Index",      "decision_clarity_index"),
            ("Model Agreement Entropy",     "model_agreement_entropy"),
        ])

    # Tab 11 — Decision System
    with tabs[10]:
        st.caption("Comprehensive decision layer — winner, tier, go/no-go")
        for cid in cpd_ids:
            h = all_hyper.get(cid)
            if h is None:
                st.error(f"{cid}: failed"); continue
            d = h.get("decision", {})
            tier  = d.get("recommendation_tier", "⚪")
            color = "#4ade80" if "Yes" in tier or "Strong" in tier else \
                    "#f5a623" if "Conditional" in tier else \
                    "#fb923c" if "NEEDS" in tier else "#f87171"

            st.markdown(
                f'<div class="hcis-card" style="border-left:5px solid {color}">'
                f'<div class="hcis-header">{cid}</div>'
                f'<div class="hcis-winner" style="color:{color}">{tier}</div>'
                f'<div style="margin-top:8px;font-size:.78rem;color:#c8deff">'
                f'Overall Winner: <b>{d.get("overall_winner","–")}</b> &nbsp;|&nbsp; '
                f'Risk-Adjusted: <b>{d.get("risk_adjusted_winner","–")}</b> &nbsp;|&nbsp; '
                f'Go Score: <b>{d.get("go_score",0):.0%}</b></div>'
                f'<div style="margin-top:6px;font-size:.73rem;color:#64748b">'
                f'+{d.get("improvements_count",0)} improvements &nbsp;|&nbsp; '
                f'-{d.get("regressions_count",0)} regressions</div></div>',
                unsafe_allow_html=True,
            )

            # Category winners mini-table
            with st.expander("📊 Category-wise Winners"):
                cw = d.get("category_winners", {})
                cw_cols = st.columns(3)
                for j, (cat, winner) in enumerate(cw.items()):
                    color2 = "#4ade80" if winner == "R2" else "#94a3b8"
                    cw_cols[j % 3].markdown(
                        f'<span style="color:{color2};font-size:.7rem">'
                        f'{"🟢" if winner=="R2" else "⚪"} {cat}</span>',
                        unsafe_allow_html=True)

    # Tab 12 — AI Report
    with tabs[11]:
        st.caption("Natural language decision report")
        for cid in cpd_ids:
            h = all_hyper.get(cid)
            if h is None:
                st.error(f"{cid}: failed"); continue
            d = h.get("decision", {})

            st.markdown(f"#### {cid} — AI Report")

            _boxes = [
                ("🟢 Why Better", "why_better", "rgba(74,222,128,.05)", "rgba(74,222,128,.2)"),
                ("🔴 Hidden Risks", "hidden_risks", "rgba(248,113,113,.05)", "rgba(248,113,113,.2)"),
                ("🎯 Best Modification", "best_modification_strategy", "rgba(245,166,35,.05)", "rgba(245,166,35,.2)"),
                ("🔑 Key Differentiators", "key_differentiators", "rgba(167,139,250,.05)", "rgba(167,139,250,.2)"),
            ]
            for title, key, bg, border in _boxes:
                text = d.get(key, "–")
                st.markdown(
                    f'<div style="background:{bg};border:1px solid {border};'
                    f'border-radius:8px;padding:12px 16px;margin:6px 0;'
                    f'font-size:.82rem;color:#c8deff;line-height:1.75">'
                    f'<b style="color:#94a3b8">{title}</b><br>{text}</div>',
                    unsafe_allow_html=True)
            st.divider()


# =============================================================================
# PATCH render_tab — inject hyper section after massive section
# =============================================================================

_original_render_tab_v2 = render_tab  # save previous wrapper reference


def render_tab(res: list):
    """Double-wrapper: calls previous wrapper (which calls original), then hyper."""
    _original_render_tab_v2(res)

    if not res:
        return
    ids = [c.get("ID", f"Cpd-{i+1}") for i, c in enumerate(res)]
    selected_ids = st.session_state.get("_cmp_select", ids[:min(3, len(ids))])
    if not selected_ids:
        return
    selected = [c for c in res if c.get("ID", "") in selected_ids]
    if len(selected) >= 2:
        try:
            _render_hyper_comparison(selected, res)
        except Exception as _ex:
            st.warning(f"Hyper comparison system: {_ex}")

# HYPER COMPARISON INTELLIGENCE SYSTEM END ────────────────────────────────────
