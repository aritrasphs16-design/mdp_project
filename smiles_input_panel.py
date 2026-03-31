"""
smiles_input_panel.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · New Main-Area Input Panel (PHASE 2)
• Live SMILES validation
• SMILES counter
• Quick-insert chips
• PubChem name→SMILES fetch
• Auto-hides after processing (compact topbar)
• Does NOT modify sidebar — fully additive
• Writes result into st.session_state["_sip_smiles"] if used
────────────────────────────────────────────────────────────────────────────
"""

import streamlit as st
import requests
import urllib.parse

# ── Safe RDKit import ─────────────────────────────────────────────────────
try:
    from rdkit import Chem as _Chem
    _RDKIT_OK = True
except Exception:
    _Chem = None
    _RDKIT_OK = False

# ── Quick-insert drug chips ───────────────────────────────────────────────
_CHIPS = {
    "Aspirin":      "CC(=O)Oc1ccccc1C(=O)O",
    "Caffeine":     "CN1C=NC2=C1C(=O)N(C(=O)N2C)C",
    "Ibuprofen":    "CC(C)Cc1ccc(cc1)C(C)C(=O)O",
    "Paracetamol":  "CC(=O)Nc1ccc(O)cc1",
    "Dopamine":     "NCCc1ccc(O)c(O)c1",
    "Metformin":    "CN(C)C(=N)NC(=N)N",
    "Lisinopril":   "OC(=O)[C@@H](CCCl)N",
    "Warfarin":     "OC=1C(=O)c2ccccc2OC=1CC(=O)c1ccccc1",
}

_PANEL_CSS = """
<style>
.sip-topbar{
  background:var(--bg2,#1a1c23);
  border:1px solid rgba(232,160,32,.2);
  border-radius:10px;
  padding:10px 18px;
  display:flex;align-items:center;gap:12px;
  font-family:'JetBrains Mono',monospace;
  font-size:.62rem;color:rgba(245,166,35,.7);
  margin-bottom:12px;
}
.sip-chip{
  display:inline-block;
  background:rgba(232,160,32,.08);
  border:1px solid rgba(232,160,32,.25);
  border-radius:6px;
  padding:2px 9px;
  font-size:.6rem;
  font-family:'JetBrains Mono',monospace;
  color:rgba(245,166,35,.8);
  cursor:pointer;margin:2px;
}
.sip-valid{color:#4ade80;font-family:'JetBrains Mono',monospace;font-size:.63rem;}
.sip-invalid{color:#f87171;font-family:'JetBrains Mono',monospace;font-size:.63rem;}
</style>
"""

def _validate_smiles(text: str) -> tuple:
    """Returns (valid_count, invalid_list, total)."""
    parts = [s.strip() for s in text.split(",") if s.strip()]
    valid, invalid = 0, []
    if _RDKIT_OK and _Chem:
        for smi in parts:
            mol = _Chem.MolFromSmiles(smi)
            if mol:
                valid += 1
            else:
                invalid.append(smi[:24])
    else:
        valid = len(parts)  # can't validate without rdkit
    return valid, invalid, len(parts)

def _fetch_pubchem(name: str) -> str | None:
    """Fetch canonical SMILES from PubChem by compound name."""
    try:
        enc = urllib.parse.quote(name.strip())
        url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{enc}/property/CanonicalSMILES/JSON"
        r = requests.get(url, timeout=6)
        if r.status_code == 200:
            data = r.json()
            return data["PropertyTable"]["Properties"][0]["CanonicalSMILES"]
    except Exception:
        pass
    return None


def render_input_panel(current_input: str = "") -> str:
    """
    Render the main-area SMILES input panel.

    Returns the SMILES string to use (may be same as current_input if user
    hasn't interacted, or a new value from chip/PubChem fetch).

    IMPORTANT: This does NOT replace the sidebar input — it is supplemental.
    If the user has already entered SMILES via sidebar, current_input will
    be non-empty and the panel collapses to a compact topbar.
    """
    st.markdown(_PANEL_CSS, unsafe_allow_html=True)

    # ── Session state keys ────────────────────────────────────────────────
    if "_sip_visible" not in st.session_state:
        st.session_state["_sip_visible"] = True
    if "_sip_smiles" not in st.session_state:
        st.session_state["_sip_smiles"] = ""
    if "_sip_pubchem_result" not in st.session_state:
        st.session_state["_sip_pubchem_result"] = ""

    # Use sidebar input as base if available
    working_smiles = current_input or st.session_state.get("_sip_smiles", "")

    # ── Compact topbar when input is already present ──────────────────────
    if working_smiles.strip() and not st.session_state["_sip_visible"]:
        parts = [s.strip() for s in working_smiles.split(",") if s.strip()]
        st.markdown(
            f'<div class="sip-topbar">⬡ ChemoFilter &nbsp;|&nbsp; '
            f'{len(parts)} compound(s) loaded &nbsp;|&nbsp; '
            f'<span style="color:rgba(100,200,100,.7)">Analysis active</span></div>',
            unsafe_allow_html=True,
        )
        if st.button("✎ Edit Compound Input", key="_sip_reopen", help="Re-open the compound SMILES input panel"):
            st.session_state["_sip_visible"] = True
            st.rerun()
        return working_smiles

    # ── Full input panel ──────────────────────────────────────────────────
    with st.expander("⬡  Compound SMILES Input — Quick Entry, Live Validation & PubChem Lookup", expanded=st.session_state["_sip_visible"]):

        # Quick-insert chips row
        st.markdown("**Reference Compound Quick Insert:**", help="Click to append this compound's SMILES")
        chip_cols = st.columns(len(_CHIPS))
        for i, (name, smi) in enumerate(_CHIPS.items()):
            with chip_cols[i]:
                if st.button(name, key=f"_sip_chip_{name}", help=smi):
                    if working_smiles.strip():
                        working_smiles = working_smiles.rstrip(", ") + f", {smi}"
                    else:
                        working_smiles = smi
                    st.session_state["_sip_smiles"] = working_smiles

        st.divider()

        # PubChem name lookup
        pc_col1, pc_col2 = st.columns([3, 1])
        with pc_col1:
            pc_name = st.text_input(
                "🔍 Search by Drug or Compound Name (PubChem Database)",
                placeholder="e.g. aspirin, metformin, ibuprofen, sildenafil",
                key="_sip_pc_name",
            )
        with pc_col2:
            st.write("")
            st.write("")
            if st.button("Retrieve SMILES from PubChem", key="_sip_pc_fetch"):
                if pc_name.strip():
                    with st.spinner("Querying PubChem compound database..."):
                        result = _fetch_pubchem(pc_name)
                    if result:
                        st.session_state["_sip_pubchem_result"] = result
                        st.success(f"Found: `{result[:50]}`")
                        if working_smiles.strip():
                            working_smiles = working_smiles.rstrip(", ") + f", {result}"
                        else:
                            working_smiles = result
                        st.session_state["_sip_smiles"] = working_smiles
                    else:
                        st.error("Not found in PubChem.")

        st.divider()

        # Main SMILES text area with live validation
        new_smiles = st.text_area(
            "SMILES Input Strings (comma-separated)",
            value=working_smiles,
            height=100,
            key="_sip_textarea",
            placeholder="CC(=O)Oc1ccccc1C(=O)O, CN1C=NC2=C1C(=O)N(C(=O)N2C)C",
            help="Enter one or more canonical SMILES strings, separated by commas. Each string represents one compound.",
        )

        # Live validation feedback
        if new_smiles.strip():
            v, inv, total = _validate_smiles(new_smiles)
            vcol1, vcol2, vcol3 = st.columns(3)
            with vcol1:
                st.markdown(f'<div class="sip-valid">✓ {v}/{total} valid</div>', unsafe_allow_html=True)
            with vcol2:
                if inv:
                    st.markdown(
                        f'<div class="sip-invalid">✗ {len(inv)} invalid: {", ".join(inv[:3])}</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown('<div class="sip-valid">All structures OK</div>', unsafe_allow_html=True)
            with vcol3:
                st.markdown(
                    f'<div class="sip-valid">{total} compound(s) entered</div>',
                    unsafe_allow_html=True,
                )

        # Action buttons
        btn1, btn2, _ = st.columns([3, 3, 1])
        with btn1:
            if st.button("✓ Confirm & Analyse Compounds", key="_sip_apply", type="primary"):
                st.session_state["_sip_smiles"] = new_smiles
                st.session_state["_sip_visible"] = False
                working_smiles = new_smiles
                st.rerun()
        with btn2:
            if st.button("✕ Clear All Compounds", key="_sip_clear"):
                st.session_state["_sip_smiles"] = ""
                working_smiles = ""

    return working_smiles or current_input
