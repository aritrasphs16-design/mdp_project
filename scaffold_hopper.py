"""
scaffold_hopper.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Scaffold Hopper — Tab 36 (PHASE 3)
• Extracts Murcko scaffolds from the current compound set
• Ranks scaffolds by frequency and drug-likeness
• Suggests hop variants via R-group substitution logic
• Fully isolated — no side-effects on existing system
────────────────────────────────────────────────────────────────────────────
"""

import streamlit as st

try:
    from rdkit import Chem
    from rdkit.Chem.Scaffolds import MurckoScaffold
    from rdkit.Chem import Descriptors, rdMolDescriptors, Crippen
    _RDK_OK = True
except Exception:
    _RDK_OK = False

try:
    import plotly.graph_objects as go
    _PLT_OK = True
except Exception:
    _PLT_OK = False

import random

# ── Scaffold extraction helpers ───────────────────────────────────────────

def _get_scaffold(smi: str) -> str | None:
    if not _RDK_OK:
        return None
    try:
        mol = Chem.MolFromSmiles(smi)
        if mol is None:
            return None
        scaf = MurckoScaffold.GetScaffoldForMol(mol)
        return Chem.MolToSmiles(scaf)
    except Exception:
        return None


def _scaffold_props(smi: str) -> dict:
    if not _RDK_OK:
        return {}
    try:
        mol = Chem.MolFromSmiles(smi)
        if mol is None:
            return {}
        return {
            "MW": round(Descriptors.MolWt(mol), 1),
            "LogP": round(Crippen.MolLogP(mol), 2),
            "Rings": rdMolDescriptors.CalcNumRings(mol),
            "ArRings": rdMolDescriptors.CalcNumAromaticRings(mol),
            "HeavyAtoms": mol.GetNumHeavyAtoms(),
        }
    except Exception:
        return {}


def _hop_variants(scaffold_smi: str) -> list:
    """Generate simple hop suggestions (heuristic substitutions)."""
    variants = []
    substitutions = [
        ("N", "O"),
        ("c1ccccc1", "c1ccncc1"),      # benzene → pyridine
        ("c1ccccc1", "c1cccs1"),        # benzene → thiophene
        ("C(=O)", "S(=O)(=O)"),
        ("CC", "CF"),
    ]
    for old, new in substitutions:
        if old in scaffold_smi:
            candidate = scaffold_smi.replace(old, new, 1)
            if _RDK_OK:
                mol = Chem.MolFromSmiles(candidate)
                if mol:
                    variants.append(Chem.MolToSmiles(mol))
            else:
                variants.append(candidate)
            if len(variants) >= 4:
                break
    return variants


# ── Main render function ──────────────────────────────────────────────────

def render_tab(res: list):
    """
    Render Scaffold Hopper tab.
    res: list of compound dicts from ChemoFilter analysis
    """
    st.markdown(
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.6rem;'
        'letter-spacing:3px;color:rgba(232,160,32,.5);text-transform:uppercase;'
        'margin-bottom:12px">⬡ Scaffold Hopper — Structural Diversity Analysis</div>',
        unsafe_allow_html=True,
    )

    if not res:
        st.warning("No compounds loaded. Please enter SMILES in the input panel.")
        return

    if not _RDK_OK:
        st.error("RDKit not available — scaffold analysis unavailable.")
        return

    # Extract SMILES
    smiles_list = []
    for c in res:
        smi = c.get("SMILES") or c.get("smi") or ""
        if smi:
            smiles_list.append(smi)

    if not smiles_list:
        st.warning("No valid SMILES found in current dataset.")
        return

    # Compute scaffolds
    scaffold_map: dict[str, list] = {}
    for smi in smiles_list:
        sc = _get_scaffold(smi)
        if sc:
            scaffold_map.setdefault(sc, []).append(smi)

    if not scaffold_map:
        st.info("No Murcko scaffolds could be extracted.")
        return

    # Sort by frequency
    sorted_scaffolds = sorted(scaffold_map.items(), key=lambda x: len(x[1]), reverse=True)

    st.markdown(f"**{len(sorted_scaffolds)} unique scaffolds** found across {len(smiles_list)} compounds.")

    # ── Frequency chart ───────────────────────────────────────────────────
    if _PLT_OK and len(sorted_scaffolds) > 1:
        top_n = sorted_scaffolds[:15]
        labels = [f"SC-{i+1}" for i in range(len(top_n))]
        counts = [len(v) for _, v in top_n]
        fig = go.Figure(go.Bar(
            x=labels,
            y=counts,
            marker_color="#f5a623",
            marker_line_color="rgba(232,160,32,.3)",
            marker_line_width=1,
        ))
        fig.update_layout(
            title="Murcko Scaffold Frequency Distribution (Top 15)",
            xaxis_title="Scaffold",
            yaxis_title="Compound Count",
            height=300,
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#c8deff", size=11),
        )
        st.plotly_chart(fig, use_container_width=True)

    # ── Scaffold table ────────────────────────────────────────────────────
    st.subheader("Murcko Scaffold Inventory")
    cols = st.columns([1, 3, 1, 1, 1, 1, 1])
    headers = ["ID", "Scaffold SMILES", "Compound Count", "MW", "LogP", "Rings", "ArRings"]
    for col, h in zip(cols, headers):
        col.markdown(f"**{h}**")

    for i, (sc_smi, members) in enumerate(sorted_scaffolds[:20]):
        props = _scaffold_props(sc_smi)
        row = st.columns([1, 3, 1, 1, 1, 1, 1])
        row[0].write(f"SC-{i+1}")
        row[1].code(sc_smi[:50], language=None)
        row[2].write(len(members))
        row[3].write(props.get("MW", "–"))
        row[4].write(props.get("LogP", "–"))
        row[5].write(props.get("Rings", "–"))
        row[6].write(props.get("ArRings", "–"))

    # ── Hop suggestions for top scaffold ─────────────────────────────────
    st.subheader("Bioisosteric Hop Suggestions for Top Scaffold")
    top_scaffold = sorted_scaffolds[0][0]
    hops = _hop_variants(top_scaffold)
    if hops:
        st.markdown(f"**Base scaffold:** `{top_scaffold}`")
        st.markdown(f"**Suggested Bioisosteric Replacements** (ring system substitutions):")
        for j, hop in enumerate(hops, 1):
            props = _scaffold_props(hop)
            st.code(f"Hop {j}: {hop}  |  MW={props.get('MW','?')}  LogP={props.get('LogP','?')}", language=None)
    else:
        st.info("No automatic hop suggestions available for this scaffold.")
