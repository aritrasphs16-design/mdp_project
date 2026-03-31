"""
reaction_simulator.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Reaction Simulator — Tab 39 (PHASE 3)
• Applies common medicinal chemistry transformations to input compounds
• Ester hydrolysis, amide formation, N-methylation, halogenation, etc.
• Shows before/after property delta
• Fully isolated — only RDKit AllChem used
────────────────────────────────────────────────────────────────────────────
"""

import streamlit as st

try:
    from rdkit import Chem
    from rdkit.Chem import AllChem, Descriptors, Crippen, rdMolDescriptors
    from rdkit.Chem import rdChemReactions
    _RDK_OK = True
except Exception:
    _RDK_OK = False

try:
    import plotly.graph_objects as go
    _PLT_OK = True
except Exception:
    _PLT_OK = False

# ── Reaction SMARTS definitions ───────────────────────────────────────────
_REACTIONS = {
    "Ester Hydrolysis":      "[C:1](=[O:2])[O:3][C:4]>>[C:1](=[O:2])[OH]",
    "N-Methylation":         "[NH2:1]>>[NH:1]C",
    "Amide Formation":       "[C:1](=[O:2])[OH:3]>>[C:1](=[O:2])N",
    "Para-Fluorination":     "[c:1][H]>>[c:1]F",
    "Ester Formation":       "[C:1](=[O:2])[OH]>>[C:1](=[O:2])OCC",
    "Hydroxylation":         "[CH3:1]>>[CH2:1]O",
}

_REACTION_COLORS = {
    "Ester Hydrolysis":  "#38bdf8",
    "N-Methylation":     "#a78bfa",
    "Amide Formation":   "#4ade80",
    "Para-Fluorination": "#f5a623",
    "Ester Formation":   "#fb923c",
    "Hydroxylation":     "#f87171",
}


def _get_props(smi: str) -> dict:
    if not _RDK_OK:
        return {}
    try:
        mol = Chem.MolFromSmiles(smi)
        if mol is None:
            return {}
        return {
            "MW":   round(Descriptors.MolWt(mol), 1),
            "LogP": round(Crippen.MolLogP(mol), 2),
            "TPSA": round(Descriptors.TPSA(mol), 1),
            "HBD":  rdMolDescriptors.CalcNumHBD(mol),
            "HBA":  rdMolDescriptors.CalcNumHBA(mol),
            "QED":  round(Descriptors.qed(mol), 3),
        }
    except Exception:
        return {}


def _apply_reaction(rxn_smarts: str, smi: str) -> list:
    """Apply SMARTS reaction to SMILES. Returns list of product SMILES."""
    if not _RDK_OK:
        return []
    try:
        rxn = rdChemReactions.ReactionFromSmarts(rxn_smarts)
        mol = Chem.MolFromSmiles(smi)
        if rxn is None or mol is None:
            return []
        products = rxn.RunReactants((mol,))
        result = []
        seen = set()
        for prod_tuple in products:
            for prod in prod_tuple:
                try:
                    Chem.SanitizeMol(prod)
                    ps = Chem.MolToSmiles(prod)
                    if ps and ps not in seen:
                        seen.add(ps)
                        result.append(ps)
                except Exception:
                    pass
        return result[:3]  # max 3 products
    except Exception:
        return []


# ── Main render function ──────────────────────────────────────────────────

def render_tab(res: list):
    st.markdown(
        '<div style="font-family:\'JetBrains Mono\',monospace;font-size:.6rem;'
        'letter-spacing:3px;color:rgba(232,160,32,.5);text-transform:uppercase;'
        'margin-bottom:12px">⬡ Reaction Simulator — Medicinal Chemistry Transformations</div>',
        unsafe_allow_html=True,
    )

    if not res:
        st.warning("No compounds loaded.")
        return

    if not _RDK_OK:
        st.error("RDKit not available — reaction simulation unavailable.")
        return

    # Compound picker
    ids = [c.get("ID", f"Cpd-{i+1}") for i, c in enumerate(res)]
    sel_id = st.selectbox("Select Compound for Reaction Simulation", ids, key="_rxnsim_sel")
    compound = next((c for c in res if c.get("ID", "") == sel_id), res[0])
    base_smi = compound.get("SMILES") or compound.get("smi") or ""

    if not base_smi:
        st.warning("Selected compound has no SMILES.")
        return

    st.markdown(f"**Base SMILES:** `{base_smi}`")
    base_props = _get_props(base_smi)
    if base_props:
        pcols = st.columns(len(base_props))
        for col, (k, v) in zip(pcols, base_props.items()):
            col.metric(k, v)

    st.divider()

    # Reaction selector
    rxn_name = st.selectbox(
        "Select Medicinal Chemistry Transformation",
        list(_REACTIONS.keys()),
        key="_rxnsim_rxn",
    )
    rxn_smarts = _REACTIONS[rxn_name]

    if st.button("▶ Apply Transformation", key="_rxnsim_run", type="primary"):
        with st.spinner(f"Applying {rxn_name}..."):
            products = _apply_reaction(rxn_smarts, base_smi)

        if products:
            st.success(f"{len(products)} product(s) generated")
            color = _REACTION_COLORS.get(rxn_name, "#f5a623")

            for i, prod_smi in enumerate(products, 1):
                prod_props = _get_props(prod_smi)
                st.markdown(
                    f'<div style="border:1px solid {color};border-radius:8px;'
                    f'padding:12px;margin:8px 0;">'
                    f'<b>Product {i}</b>: <code>{prod_smi}</code></div>',
                    unsafe_allow_html=True,
                )

                if base_props and prod_props:
                    # Delta display
                    delta_cols = st.columns(len(base_props))
                    for col, (k, base_v) in zip(delta_cols, base_props.items()):
                        prod_v = prod_props.get(k, base_v)
                        try:
                            delta = round(float(prod_v) - float(base_v), 2)
                            sign = "+" if delta >= 0 else ""
                            d_color = "#4ade80" if delta >= 0 else "#f87171"
                            col.markdown(
                                f"**{k}** `{prod_v}`  "
                                f'<span style="color:{d_color};font-size:.75rem">({sign}{delta})</span>',
                                unsafe_allow_html=True,
                            )
                        except Exception:
                            col.write(f"{k}: {prod_v}")

                st.divider()
        else:
            st.info(
                f"No products generated for **{rxn_name}** on this compound. "
                "The reaction pattern may not match any substructure."
            )

    # ── All reactions summary ─────────────────────────────────────────────
    with st.expander("Show All Reaction Outcomes (Batch Simulation)"):
        for rname, rsmarts in _REACTIONS.items():
            prods = _apply_reaction(rsmarts, base_smi)
            status = f"✓ {len(prods)} product(s)" if prods else "✗ No match"
            color  = "#4ade80" if prods else "#94a3b8"
            st.markdown(
                f'<span style="color:{color};font-family:JetBrains Mono,monospace;'
                f'font-size:.7rem">{rname}: {status}</span>',
                unsafe_allow_html=True,
            )
