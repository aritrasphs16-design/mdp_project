"""
drug_class_predictor.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Drug Class Predictor — Tab 38 (PHASE 3)
• Rule-based classification of compounds into drug categories
• Uses physicochemical properties + substructure patterns
• Categories: CNS, Cardiovascular, Anti-infective, Metabolic, Oncology, Other
• Fully isolated — no external dependencies beyond RDKit
────────────────────────────────────────────────────────────────────────────
"""

import streamlit as st

try:
    from rdkit import Chem
    from rdkit.Chem import Descriptors, rdMolDescriptors, Crippen
    _RDK_OK = True
except Exception:
    _RDK_OK = False

try:
    import plotly.graph_objects as go
    _PLT_OK = True
except Exception:
    _PLT_OK = False

# ── Classification rules (SMARTS + property thresholds) ──────────────────

_CNS_SMARTS = [
    "c1ccc2[nH]ccc2c1",         # indole
    "c1cnc2ccccc2c1",            # quinoline-like
    "c1ccncc1",                  # pyridine ring
    "N1CCNCC1",                  # piperazine
]

_CARDIO_SMARTS = [
    "C(=O)[OH]",                 # carboxylic acid
    "c1ccc(S(=O)(=O)N)cc1",     # sulfonamide on ring
    "C1=CC=C(C=C1)Cl",          # chlorobenzene motif
]

_ANTIINFECTIVE_SMARTS = [
    "c1nc2ccccc2[nH]1",         # benzimidazole
    "c1ccc(-c2nc3ccccc3s2)cc1", # benzothiazole
    "[NH2]c1ccc(cc1)S(=O)(=O)N",# sulfanilamide
    "c1cnc(N)nc1",              # aminopyrimidine
]

_ONCOLOGY_SMARTS = [
    "c1ccc2nc3ccccc3cc2c1",     # acridine
    "N#Cc1ccccc1",              # benzonitrile
    "c1ccc(Cl)cc1",             # chlorophenyl
]

_METABOLIC_SMARTS = [
    "C(=O)N",                   # amide
    "OCC(O)CO",                 # polyol / sugar-like
    "[NH]C(=N)N",               # guanidine
]


def _match_smarts(mol, smarts_list: list) -> int:
    if not _RDK_OK or mol is None:
        return 0
    count = 0
    for sma in smarts_list:
        try:
            pat = Chem.MolFromSmarts(sma)
            if pat and mol.GetSubstructMatches(pat):
                count += 1
        except Exception:
            pass
    return count


def _classify(compound: dict) -> tuple:
    """Returns (class_label, confidence_pct, reasoning)."""
    if not _RDK_OK:
        return "Unknown", 50, "RDKit unavailable"

    smi = compound.get("SMILES") or compound.get("smi") or ""
    mol = None
    if smi:
        try:
            mol = Chem.MolFromSmiles(smi)
        except Exception:
            pass

    mw   = float(compound.get("MW", 400))
    logp = float(compound.get("LogP", 2.0))
    tpsa = float(compound.get("tPSA", compound.get("TPSA", 80)))

    scores = {
        "CNS":              0,
        "Cardiovascular":   0,
        "Anti-infective":   0,
        "Oncology":         0,
        "Metabolic/Other":  0,
    }

    # CNS: low MW, low TPSA, moderate LogP, CNS-like rings
    if mw < 450 and tpsa < 90 and 1 <= logp <= 5:
        scores["CNS"] += 3
    if mol:
        scores["CNS"]             += _match_smarts(mol, _CNS_SMARTS) * 2
        scores["Cardiovascular"]  += _match_smarts(mol, _CARDIO_SMARTS) * 2
        scores["Anti-infective"]  += _match_smarts(mol, _ANTIINFECTIVE_SMARTS) * 2
        scores["Oncology"]        += _match_smarts(mol, _ONCOLOGY_SMARTS) * 2
        scores["Metabolic/Other"] += _match_smarts(mol, _METABOLIC_SMARTS) * 2

    # Property-based boosts
    if tpsa > 140:
        scores["Metabolic/Other"] += 2
    if logp > 4.5:
        scores["Oncology"] += 1
    if mw > 500:
        scores["Metabolic/Other"] += 1

    best_class = max(scores, key=scores.get)
    total = sum(scores.values()) or 1
    confidence = min(95, max(30, int(scores[best_class] / total * 100)))
    reasoning = (
        f"MW={mw}, LogP={logp}, TPSA={tpsa}; "
        f"top pattern matches: {scores[best_class]}/{total}"
    )
    return best_class, confidence, reasoning


# ── Main render function ──────────────────────────────────────────────────

def render_tab(res: list):
    st.markdown(
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.6rem;'
        'letter-spacing:3px;color:rgba(232,160,32,.5);text-transform:uppercase;'
        'margin-bottom:12px">⬡ Drug Class Predictor — Rule-Based Classification</div>',
        unsafe_allow_html=True,
    )

    if not res:
        st.warning("No compounds loaded.")
        return

    # Classify all compounds
    class_counts: dict[str, int] = {}
    table_rows = []

    for cpd in res[:50]:  # cap for performance
        cls, conf, reason = _classify(cpd)
        class_counts[cls] = class_counts.get(cls, 0) + 1
        table_rows.append({
            "ID": cpd.get("ID", "–"),
            "Class": cls,
            "Confidence": f"{conf}%",
            "MW": cpd.get("MW", "–"),
            "LogP": cpd.get("LogP", "–"),
            "Classification Basis": reason,
        })

    # ── Distribution pie chart ────────────────────────────────────────────
    if _PLT_OK and class_counts:
        colors = ["#f5a623", "#4ade80", "#38bdf8", "#a78bfa", "#fb923c", "#f87171"]
        fig = go.Figure(go.Pie(
            labels=list(class_counts.keys()),
            values=list(class_counts.values()),
            marker_colors=colors[:len(class_counts)],
            hole=0.45,
            textinfo="label+percent",
            textfont_size=11,
        ))
        fig.update_layout(
            title="Predicted Therapeutic Drug Class Distribution",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#c8deff"),
            height=340,
            showlegend=True,
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── Per-compound table ────────────────────────────────────────────────
    st.subheader("Per-Compound Therapeutic Class Prediction")

    _CLASS_COLORS = {
        "CNS":              "#a78bfa",
        "Cardiovascular":   "#f87171",
        "Anti-infective":   "#4ade80",
        "Oncology":         "#fb923c",
        "Metabolic/Other":  "#38bdf8",
        "Unknown":          "#94a3b8",
    }

    hdr = st.columns([1, 2, 1, 1, 1, 3])
    for col, h in zip(hdr, ["ID", "Class", "Confidence", "MW", "LogP", "Classification Basis"]):
        col.markdown(f"**{h}**")

    for row in table_rows:
        r = st.columns([1, 2, 1, 1, 1, 3])
        r[0].write(row["ID"])
        color = _CLASS_COLORS.get(row["Class"], "#94a3b8")
        r[1].markdown(
            f'<span style="color:{color};font-weight:600">{row["Class"]}</span>',
            unsafe_allow_html=True,
        )
        r[2].write(row["Confidence"])
        r[3].write(row["MW"])
        r[4].write(row["LogP"])
        r[5].write(row["Classification Basis"][:60])
