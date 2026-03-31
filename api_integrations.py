"""
api_integrations.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · External Scientific Intelligence — UI Layer
Tier 1: Core  (PubChem, ChEMBL, PDB)
Tier 2: Extended  (OpenFDA, UniProt, KEGG, BindingDB, Open Targets …)
Tier 3: Experimental  (Literature, Genomics, Pharmacology …)
All calls: lazy, cached, fail-safe. App works 100% offline.
────────────────────────────────────────────────────────────────────────────
"""
import streamlit as st

try:
    from api_registry import API_REGISTRY, TIER_META, TIER_ORDER, TIER_CORE, TIER_EXTENDED, TIER_EXPERIMENTAL, count_by_tier
    _REG_OK = True
except Exception:
    _REG_OK = False
    API_REGISTRY = {}; TIER_META = {}; TIER_ORDER = ["core","extended","experimental"]
    TIER_CORE = {}; TIER_EXTENDED = {}; TIER_EXPERIMENTAL = {}
    def count_by_tier(): return {}

try:
    from api_manager import fetch_api, is_enabled, set_enabled, status_badge, render_api_status_row
    _MGR_OK = True
except Exception:
    _MGR_OK = False
    def fetch_api(*a, **k): return {"status":"failed","data":{},"error":"api_manager unavailable"}
    def is_enabled(k): return False
    def set_enabled(k, v): pass
    def status_badge(r): return "⚪"
    def render_api_status_row(*a): pass

# Keep old pubchem functions for backward compatibility
try:
    from api_manager import _fetch_pubchem as fetch_pubchem_by_smiles
except Exception:
    def fetch_pubchem_by_smiles(s): return {}

def fetch_pubchem_by_name(name: str) -> dict:
    try:
        import urllib.parse, requests as _r
        enc = urllib.parse.quote(name.strip())
        url = f"https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/{enc}/property/CanonicalSMILES/JSON"
        r = _r.get(url, timeout=5); r.raise_for_status()
        smi = r.json().get("PropertyTable",{}).get("Properties",[{}])[0].get("CanonicalSMILES","")
        if smi:
            res = fetch_pubchem_by_smiles(smi)
            return res.get("data", {}) if isinstance(res, dict) and "data" in res else res
        return {"error": "SMILES not found"}
    except Exception as e:
        return {"error": str(e)[:100]}

# ── Styling constants ─────────────────────────────────────────────────────
_TIER_COLORS = {"core": "#4ade80", "extended": "#f5a623", "experimental": "#a78bfa"}
_TIER_BG     = {"core": "rgba(74,222,128,.06)", "extended": "rgba(245,166,35,.06)", "experimental": "rgba(167,139,250,.06)"}

def _card(content: str, color: str = "#f5a623", border_left: bool = True) -> str:
    border = f"border-left:4px solid {color};" if border_left else ""
    return (f'<div style="background:rgba(14,16,23,.6);{border}'
            f'border:1px solid {color}25;border-radius:8px;'
            f'padding:12px 16px;margin:6px 0;font-size:.78rem;color:#c8deff;line-height:1.7">'
            f'{content}</div>')

def _kv(key: str, val, mono: bool = True) -> str:
    v = f'<code style="color:#f5a623;background:rgba(245,166,35,.08);padding:1px 6px;border-radius:3px">{val}</code>' if mono else f'<span style="color:#c8deff">{val}</span>'
    return f'<div style="display:flex;gap:8px;padding:3px 0;border-bottom:1px solid rgba(255,255,255,.03)"><span style="color:#64748b;font-size:.72rem;min-width:160px">{key}</span>{v}</div>'

# ═══════════════════════════════════════════════════════════════════════════
# MAIN RENDER FUNCTION — External Scientific Intelligence
# ═══════════════════════════════════════════════════════════════════════════

def render_external_data_section(compound: dict):
    """
    Renders the full tiered External Scientific Intelligence section.
    All API calls are lazy — only fire on button click.
    100% safe if all APIs fail.
    """
    smiles   = compound.get("SMILES") or compound.get("smi") or ""
    cpd_id   = compound.get("ID", "Compound")
    cpd_name = compound.get("name", cpd_id)

    st.markdown("---")
    st.markdown("## 🌐 External Scientific Intelligence")

    # Summary header
    counts = count_by_tier()
    st.markdown(
        f'<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px">'
        f'<span style="background:rgba(74,222,128,.08);border:1px solid rgba(74,222,128,.25);'
        f'border-radius:6px;padding:4px 12px;font-size:.7rem;color:#4ade80">'
        f'🟢 {counts.get("core",3)} Core APIs</span>'
        f'<span style="background:rgba(245,166,35,.08);border:1px solid rgba(245,166,35,.25);'
        f'border-radius:6px;padding:4px 12px;font-size:.7rem;color:#f5a623">'
        f'🟡 {counts.get("extended",28)} Extended APIs</span>'
        f'<span style="background:rgba(167,139,250,.08);border:1px solid rgba(167,139,250,.25);'
        f'border-radius:6px;padding:4px 12px;font-size:.7rem;color:#a78bfa">'
        f'🟣 {counts.get("experimental",20)} Experimental APIs</span>'
        f'<span style="color:#475569;font-size:.68rem;padding:4px 0">'
        f'All lazy-loaded · 24h cached · Offline-safe</span></div>',
        unsafe_allow_html=True,
    )

    if not smiles:
        st.info("No SMILES available — cannot query external databases for this compound.")
        return

    # ── LAYER 1: CORE ──────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:{_TIER_BG["core"]};border:1px solid rgba(74,222,128,.2);'
        f'border-radius:10px;padding:14px 18px;margin:10px 0">'
        f'<b style="color:#4ade80">🟢 Layer 1 — Core Databases</b>'
        f'<span style="color:#64748b;font-size:.72rem;margin-left:12px">'
        f'Always-available reference databases · Cached 24h</span></div>',
        unsafe_allow_html=True)

    core_tabs = st.tabs(["🧪 PubChem", "🎯 ChEMBL Bioactivity", "🏗️ Protein Data Bank (PDB)"])

    with core_tabs[0]:
        _render_pubchem(smiles, cpd_id, compound)

    with core_tabs[1]:
        _render_chembl(smiles, cpd_id)

    with core_tabs[2]:
        _render_pdb(cpd_id)

    # ── LAYER 2: EXTENDED ──────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:{_TIER_BG["extended"]};border:1px solid rgba(245,166,35,.2);'
        f'border-radius:10px;padding:14px 18px;margin:10px 0">'
        f'<b style="color:#f5a623">🟡 Layer 2 — Extended Scientific Databases</b>'
        f'<span style="color:#64748b;font-size:.72rem;margin-left:12px">'
        f'Click to fetch · Cached 24h per compound</span></div>',
        unsafe_allow_html=True)

    ext_tabs = st.tabs([
        "🏛️ FDA Drug Labels",
        "🧬 UniProt Targets",
        "🗺️ KEGG Pathways",
        "🏥 Clinical Trials",
        "🛡️ GHS Safety",
        "💊 GtoPdb Pharmacology",
        "🌵 NCI Structure",
        "🔄 UniChem Cross-Refs",
    ])

    with ext_tabs[0]: _render_openfda(smiles, cpd_id)
    with ext_tabs[1]: _render_uniprot(cpd_id)
    with ext_tabs[2]: _render_kegg(cpd_id)
    with ext_tabs[3]: _render_clinicaltrials(cpd_id)
    with ext_tabs[4]: _render_chemrisk(smiles, cpd_id, compound)
    with ext_tabs[5]: _render_gtopdb(cpd_id)
    with ext_tabs[6]: _render_nci_cactus(cpd_id)
    with ext_tabs[7]: _render_unichem(compound)

    # ── LAYER 3: EXPERIMENTAL ──────────────────────────────────────────────
    st.markdown(
        f'<div style="background:{_TIER_BG["experimental"]};border:1px solid rgba(167,139,250,.2);'
        f'border-radius:10px;padding:14px 18px;margin:10px 0">'
        f'<b style="color:#a78bfa">🟣 Layer 3 — Experimental & Literature APIs</b>'
        f'<span style="color:#64748b;font-size:.72rem;margin-left:12px">'
        f'Opt-in only · Toggle to enable per session</span></div>',
        unsafe_allow_html=True)

    exp_tabs = st.tabs([
        "📰 Europe PMC Literature",
        "🤖 Semantic Scholar AI",
        "🏥 DisGeNET Disease",
        "📋 API Registry Browser",
    ])

    with exp_tabs[0]: _render_europe_pmc(cpd_id)
    with exp_tabs[1]: _render_semantic_scholar(cpd_id)
    with exp_tabs[2]: _render_disgenet(cpd_id)
    with exp_tabs[3]: _render_api_registry_browser()


# ═══════════════════════════════════════════════════════════════════════════
# INDIVIDUAL API PANEL RENDERERS
# ═══════════════════════════════════════════════════════════════════════════

def _fetch_btn(label: str, key: str) -> bool:
    return st.button(label, key=key, type="primary")


# ── PubChem ───────────────────────────────────────────────────────────────
def _render_pubchem(smiles: str, cpd_id: str, compound: dict):
    st.caption("NIH PubChem — Verified compound identity, physicochemical properties & synonyms")
    _ss = f"_apiv2_pubchem_{cpd_id}"
    if not st.session_state.get(_ss):
        st.markdown(_card(
            "Fetches verified molecular properties, IUPAC name, synonyms, InChIKey "
            "and structure data from the NIH PubChem compound database.",
            "#4ade80"), unsafe_allow_html=True)
        if _fetch_btn("🔍 Fetch PubChem Data", f"_btn_pc2_{cpd_id}"):
            st.session_state[_ss] = True; st.rerun()
        return
    with st.spinner("Querying PubChem..."):
        res = fetch_api("pubchem", smiles=smiles)
    if res.get("status") != "ok":
        st.warning(f"PubChem: {res.get('error','Failed')}"); return
    d = res["data"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("CID", d.get("cid","N/A"), help="PubChem Compound ID")
    c2.metric("Molecular Weight (Da)", d.get("mw","N/A"))
    c3.metric("Lipophilicity (XLogP)", d.get("xlogp","N/A"), help="XLogP from PubChem")
    c4.metric("Polar Surface Area (Ų)", d.get("tpsa","N/A"))
    st.markdown("".join([
        _kv("IUPAC Name", d.get("iupac_name","N/A")),
        _kv("Molecular Formula", d.get("formula","N/A")),
        _kv("InChIKey", d.get("inchikey","N/A")),
        _kv("Hydrogen Bond Donors", d.get("hbd","N/A")),
        _kv("Hydrogen Bond Acceptors", d.get("hba","N/A")),
        _kv("Rotatable Bonds", d.get("rotatable_bonds","N/A")),
        _kv("Structural Complexity", d.get("complexity","N/A")),
        _kv("Formal Charge", d.get("charge","N/A")),
    ]), unsafe_allow_html=True)
    if d.get("synonyms"):
        st.markdown("**Known Names / Drug Synonyms:** " + " · ".join(f"`{s}`" for s in d["synonyms"]))
    # Property cross-check vs local calculation
    loc_mw = compound.get("MW","?"); loc_lp = compound.get("LogP","?")
    with st.expander("🔍 PubChem vs Local RDKit Cross-Verification"):
        vc = st.columns(2)
        with vc[0]:
            try: dmw = round(float(d.get("mw",0)) - float(loc_mw), 2); dm_str = f"{dmw:+.2f} Da"
            except: dm_str = "N/A"
            st.metric("MW: PubChem vs RDKit", f"{d.get('mw','?')} Da", delta=dm_str, help="Difference vs local RDKit calculation")
        with vc[1]:
            try: dlp = round(float(d.get("xlogp",0)) - float(loc_lp), 2); dlp_str = f"{dlp:+.2f}"
            except: dlp_str = "N/A"
            st.metric("LogP: PubChem XLogP vs Crippen", f"{d.get('xlogp','?')}", delta=dlp_str)
    st.markdown(f'<a href="{d.get("pubchem_url","#")}" target="_blank" style="color:#4ade80;font-size:.75rem">🔗 View on PubChem →</a>', unsafe_allow_html=True)


# ── ChEMBL ────────────────────────────────────────────────────────────────
def _render_chembl(smiles: str, cpd_id: str):
    st.caption("ChEMBL (EMBL-EBI) — Experimental bioactivity, IC50/Ki values & protein targets")
    _ss = f"_apiv2_chembl_{cpd_id}"
    if not st.session_state.get(_ss):
        st.markdown(_card("Retrieves measured bioactivity data (IC₅₀, Ki), known protein targets and assay descriptions from the ChEMBL database.", "#4ade80"), unsafe_allow_html=True)
        if _fetch_btn("🎯 Fetch ChEMBL Bioactivity Data", f"_btn_ch2_{cpd_id}"):
            st.session_state[_ss] = True; st.rerun()
        return
    with st.spinner("Querying ChEMBL..."):
        res = fetch_api("chembl", smiles=smiles)
    if res.get("status") != "ok":
        st.warning(f"ChEMBL: {res.get('error','Failed')}"); return
    d = res["data"]
    st.markdown(f"**Top Structural Match:** `{d.get('mol_name','N/A')}` (ChEMBL ID: `{d.get('chembl_id','N/A')}`, Tanimoto Similarity: `{float(d.get('similarity',0)):.0f}%`)")
    if d.get("bioactivities"):
        st.markdown(f"**Experimental Bioactivity Data ({d['assay_count']} assays retrieved):**")
        import pandas as pd
        df = pd.DataFrame(d["bioactivities"])
        df.columns = ["Protein Target", "Activity Type", "Value", "Units", "Publication Year"]
        st.dataframe(df, use_container_width=True, height=220)
    else:
        st.info("No experimental bioactivity data found for structurally similar compounds.")
    st.markdown(f'<a href="{d.get("chembl_url","#")}" target="_blank" style="color:#4ade80;font-size:.75rem">🔗 View on ChEMBL →</a>', unsafe_allow_html=True)


# ── PDB ───────────────────────────────────────────────────────────────────
def _render_pdb(cpd_id: str):
    st.caption("RCSB PDB — Protein crystal structures, binding sites & co-crystal ligand context")
    _ss = f"_apiv2_pdb_{cpd_id}"
    pdb_q = st.text_input("Search PDB by target or mechanism keyword", placeholder="e.g. kinase, EGFR, protease, GPCR", key=f"_pdbq2_{cpd_id}")
    if _fetch_btn("🏗️ Search Protein Data Bank", f"_btn_pdb2_{cpd_id}"):
        if pdb_q.strip(): st.session_state[_ss] = pdb_q.strip()
    if st.session_state.get(_ss):
        with st.spinner("Searching PDB..."):
            res = fetch_api("pdb", query=st.session_state[_ss])
        if res.get("status") != "ok":
            st.warning(f"PDB: {res.get('error','Failed')}"); return
        d = res["data"]
        st.markdown(f"**{d.get('total_count',0)} total structures found.** Top results:")
        for e in d.get("entries", []):
            st.markdown(f'<div style="display:flex;gap:12px;align-items:center;padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04)"><code style="color:#f5a623;font-size:.82rem">{e["pdb_id"]}</code><a href="{e["url"]}" target="_blank" style="color:#38bdf8;font-size:.74rem">View Structure on RCSB →</a><span style="color:#475569;font-size:.65rem">Relevance: {e["score"]}</span></div>', unsafe_allow_html=True)


# ── OpenFDA ───────────────────────────────────────────────────────────────
def _render_openfda(smiles: str, cpd_id: str):
    st.caption("OpenFDA — FDA-approved drug labels, indications, warnings & adverse events")
    drug_q = st.text_input("Drug name for FDA label lookup", placeholder="e.g. aspirin, metformin, ibuprofen", key=f"_fdaql_{cpd_id}")
    if _fetch_btn("🏛️ Fetch FDA Drug Label", f"_btn_fda_{cpd_id}"):
        if drug_q.strip():
            with st.spinner("Querying FDA..."):
                res = fetch_api("openfda", query=drug_q)
            if res.get("status") != "ok":
                st.warning(f"OpenFDA: {res.get('error','Not found')}"); return
            d = res["data"]
            st.markdown("".join([
                _kv("Brand Name", d.get("brand_name","N/A")),
                _kv("Generic Name", d.get("generic_name","N/A")),
                _kv("Manufacturer", d.get("manufacturer","N/A")),
                _kv("Route of Administration", d.get("route","N/A")),
            ]), unsafe_allow_html=True)
            if d.get("indications"):
                with st.expander("📋 Indications & Usage"):
                    st.write(d["indications"])
            if d.get("warnings"):
                with st.expander("⚠️ Warnings"):
                    st.write(d["warnings"])


# ── UniProt ───────────────────────────────────────────────────────────────
def _render_uniprot(cpd_id: str):
    st.caption("UniProt — Protein sequence, function & disease associations for drug targets")
    prot_q = st.text_input("Protein or gene name", placeholder="e.g. EGFR, CDK2, PCSK9, ACE2", key=f"_uniql_{cpd_id}")
    if _fetch_btn("🧬 Fetch UniProt Entry", f"_btn_uni_{cpd_id}"):
        if prot_q.strip():
            with st.spinner("Querying UniProt..."):
                res = fetch_api("uniprot", query=prot_q)
            if res.get("status") != "ok":
                st.warning(f"UniProt: {res.get('error','Failed')}"); return
            for e in res["data"].get("entries", []):
                with st.expander(f"🧬 {e.get('name','N/A')} — {e.get('protein','N/A')}"):
                    st.markdown("".join([
                        _kv("UniProt Accession", e.get("accession","N/A")),
                        _kv("Organism", e.get("organism","N/A")),
                        _kv("Function", e.get("function","N/A"), mono=False),
                    ]), unsafe_allow_html=True)


# ── KEGG ──────────────────────────────────────────────────────────────────
def _render_kegg(cpd_id: str):
    st.caption("KEGG — Metabolic pathway mapping, enzyme interactions & drug entries")
    kegg_q = st.text_input("Compound or drug name for KEGG lookup", placeholder="e.g. aspirin, ATP, glucose", key=f"_keggq_{cpd_id}")
    if _fetch_btn("🗺️ Fetch KEGG Entry", f"_btn_kegg_{cpd_id}"):
        if kegg_q.strip():
            with st.spinner("Querying KEGG..."):
                res = fetch_api("kegg", query=kegg_q)
            if res.get("status") != "ok":
                st.warning(f"KEGG: {res.get('error','Failed')}"); return
            for c in res["data"].get("compounds", []):
                st.markdown(f'`{c.get("id","?")}` — {c.get("name","N/A")}')


# ── Clinical Trials ───────────────────────────────────────────────────────
def _render_clinicaltrials(cpd_id: str):
    st.caption("ClinicalTrials.gov — Active and completed clinical trials with phase, status & sponsor")
    ct_q = st.text_input("Search clinical trials", placeholder="e.g. ibuprofen inflammation, EGFR lung cancer", key=f"_ctq_{cpd_id}")
    if _fetch_btn("🏥 Search ClinicalTrials.gov", f"_btn_ct_{cpd_id}"):
        if ct_q.strip():
            with st.spinner("Querying ClinicalTrials.gov..."):
                res = fetch_api("clinicaltrials", query=ct_q)
            if res.get("status") != "ok":
                st.warning(f"ClinicalTrials: {res.get('error','Failed')}"); return
            d = res["data"]
            st.markdown(f"**{d.get('total',0)} trials found.** Top results:")
            import pandas as pd
            df = pd.DataFrame(d.get("trials", []))
            if not df.empty:
                df.columns = ["NCT ID","Title","Phase","Status","Condition","Sponsor"]
                st.dataframe(df, use_container_width=True, height=200)


# ── GHS Safety ───────────────────────────────────────────────────────────
def _render_chemrisk(smiles: str, cpd_id: str, compound: dict):
    st.caption("PubChem GHS Classification — Hazard codes, signal words & precautionary statements")
    _ss = f"_apiv2_ghs_{cpd_id}"
    if not st.session_state.get(_ss):
        if _fetch_btn("🛡️ Fetch GHS Safety Data", f"_btn_ghs_{cpd_id}"):
            st.session_state[_ss] = True; st.rerun()
        return
    with st.spinner("Fetching GHS data..."):
        # Get CID first via pubchem result if cached
        pc = fetch_api("pubchem", smiles=smiles)
        cid = pc.get("data", {}).get("cid", 0) if pc.get("status") == "ok" else 0
        if not cid:
            st.info("Fetch PubChem data first to obtain a CID for GHS lookup."); return
        res = fetch_api("chemrisk", cid=int(cid))
    if res.get("status") != "ok":
        st.warning(f"GHS: {res.get('error','Failed')}"); return
    ghs = res["data"].get("ghs_data", {})
    if not ghs:
        st.info("No GHS classification data available for this compound.")
    for heading, values in list(ghs.items())[:8]:
        if values:
            st.markdown(f"**{heading}:**")
            for v in values[:3]:
                if v: st.markdown(f"- {v}")


# ── GtoPdb ────────────────────────────────────────────────────────────────
def _render_gtopdb(cpd_id: str):
    st.caption("Guide to Pharmacology (GtoPdb) — IUPHAR/BPS curated ligand-receptor interactions")
    gtop_q = st.text_input("Ligand or drug name for pharmacology lookup", placeholder="e.g. aspirin, adrenaline, morphine", key=f"_gtq_{cpd_id}")
    if _fetch_btn("💊 Fetch Pharmacology Data", f"_btn_gt_{cpd_id}"):
        if gtop_q.strip():
            with st.spinner("Querying GtoPdb..."):
                res = fetch_api("gtopdb", query=gtop_q)
            if res.get("status") != "ok":
                st.warning(f"GtoPdb: {res.get('error','Failed')}"); return
            for lig in res["data"].get("ligands", []):
                with st.expander(f"💊 {lig.get('name','N/A')} (Ligand ID: {lig.get('ligand_id','')})"):
                    st.markdown("".join([
                        _kv("Type", lig.get("type","N/A")),
                        _kv("Approved Drug", "Yes" if lig.get("approved") else "No"),
                    ]), unsafe_allow_html=True)
                    st.markdown(f'<a href="{lig.get("url","#")}" target="_blank" style="color:#f5a623;font-size:.73rem">🔗 View on GtoPdb →</a>', unsafe_allow_html=True)


# ── NCI CACTUS ────────────────────────────────────────────────────────────
def _render_nci_cactus(cpd_id: str):
    st.caption("NCI CACTUS — Chemical structure resolver: convert drug names to canonical SMILES")
    nci_q = st.text_input("Enter drug or compound name", placeholder="e.g. paclitaxel, tamoxifen, cisplatin", key=f"_nciq_{cpd_id}")
    if _fetch_btn("🌵 Resolve Structure via NCI", f"_btn_nci_{cpd_id}"):
        if nci_q.strip():
            with st.spinner("Querying NCI CACTUS..."):
                res = fetch_api("nci_cactus", query=nci_q)
            if res.get("status") != "ok":
                st.warning(f"NCI CACTUS: {res.get('error','Failed')}"); return
            smi = res["data"].get("smiles","")
            if smi:
                st.success(f"**Canonical SMILES:** `{smi}`")
                st.markdown("Copy this SMILES and paste into the compound input panel to analyse.")
            else:
                st.info("No SMILES returned from NCI CACTUS.")


# ── UniChem ───────────────────────────────────────────────────────────────
def _render_unichem(compound: dict):
    cpd_id  = compound.get("ID","?")
    inchikey = compound.get("InChIKey","")
    st.caption("UniChem (EMBL-EBI) — Cross-database chemical identifier mapping across 40+ databases")
    if not inchikey:
        st.info("InChIKey not available for this compound. Fetch PubChem data first.")
        return
    if _fetch_btn("🔄 Fetch Cross-Database References", f"_btn_uni2_{cpd_id}"):
        with st.spinner("Querying UniChem..."):
            res = fetch_api("unichem", inchikey=inchikey)
        if res.get("status") != "ok":
            st.warning(f"UniChem: {res.get('error','Failed')}"); return
        refs = res["data"].get("cross_references", [])
        if refs:
            import pandas as pd
            df = pd.DataFrame(refs)
            df.columns = ["Database", "Compound ID", "Database URL"]
            st.dataframe(df[["Database","Compound ID"]], use_container_width=True)
        else:
            st.info("No cross-database references found.")


# ── Europe PMC ────────────────────────────────────────────────────────────
def _render_europe_pmc(cpd_id: str):
    st.caption("Europe PMC — 37M+ open-access biomedical publications (experimental API)")
    pmc_q = st.text_input("Literature search query", placeholder="e.g. aspirin anti-inflammatory, EGFR inhibitor cancer", key=f"_pmcq_{cpd_id}")
    if _fetch_btn("📰 Search Biomedical Literature", f"_btn_pmc_{cpd_id}"):
        if pmc_q.strip():
            with st.spinner("Searching Europe PMC..."):
                res = fetch_api("europe_pmc", query=pmc_q)
            if res.get("status") != "ok":
                st.warning(f"Europe PMC: {res.get('error','Failed')}"); return
            d = res["data"]
            st.markdown(f"**{d.get('hit_count',0):,} total publications found.** Top 5:")
            for p in d.get("papers", []):
                with st.expander(f"📰 {p.get('title','N/A')[:90]}..."):
                    st.markdown("".join([
                        _kv("PubMed ID", p.get("pmid","N/A")),
                        _kv("Journal", p.get("journal","N/A")),
                        _kv("Year", p.get("year","")),
                        _kv("DOI", p.get("doi","N/A")),
                    ]), unsafe_allow_html=True)
                    if p.get("abstract"): st.markdown(f"*{p['abstract'][:400]}...*")


# ── Semantic Scholar ──────────────────────────────────────────────────────
def _render_semantic_scholar(cpd_id: str):
    st.caption("Semantic Scholar AI — AI-indexed academic papers with TL;DR summaries")
    ss_q = st.text_input("AI paper search query", placeholder="e.g. drug resistance mechanisms kinase, ADMET machine learning", key=f"_ssq_{cpd_id}")
    if _fetch_btn("🤖 Search with Semantic Scholar AI", f"_btn_ss_{cpd_id}"):
        if ss_q.strip():
            with st.spinner("Querying Semantic Scholar..."):
                res = fetch_api("semantic_scholar", query=ss_q)
            if res.get("status") != "ok":
                st.warning(f"Semantic Scholar: {res.get('error','Failed')}"); return
            for p in res["data"].get("papers", []):
                with st.expander(f"🤖 {p.get('title','N/A')[:90]}... ({p.get('year','')})"):
                    st.markdown("".join([
                        _kv("Authors", p.get("authors","N/A"), mono=False),
                        _kv("Citations", p.get("citations",0)),
                        _kv("DOI", p.get("doi","N/A")),
                    ]), unsafe_allow_html=True)
                    if p.get("tldr"): st.markdown(f"**AI Summary (TL;DR):** *{p['tldr']}*")


# ── DisGeNET ─────────────────────────────────────────────────────────────
def _render_disgenet(cpd_id: str):
    st.caption("DisGeNET — Curated gene–disease association scores from genetics, literature & clinical data")
    dg_q = st.text_input("Gene symbol", placeholder="e.g. EGFR, BRCA1, TP53, ACE2", key=f"_dgq_{cpd_id}")
    if _fetch_btn("🏥 Fetch Disease Associations", f"_btn_dg_{cpd_id}"):
        if dg_q.strip():
            with st.spinner("Querying DisGeNET..."):
                res = fetch_api("disgenet", query=dg_q)
            if res.get("status") != "ok":
                st.warning(f"DisGeNET: {res.get('error','Failed')}"); return
            assocs = res["data"].get("associations", [])
            if assocs:
                import pandas as pd
                df = pd.DataFrame(assocs)
                df.columns = ["Disease Name", "Association Score", "Disease ID"]
                st.dataframe(df, use_container_width=True)
            else:
                st.info("No disease associations found for this gene.")


# ── API Registry Browser ─────────────────────────────────────────────────
def _render_api_registry_browser():
    st.caption("Browse all 53 scientific APIs available in ChemoFilter's integration system")
    if not _REG_OK:
        st.info("API registry not available.")
        return
    tier_filter = st.selectbox("Filter by tier", ["All"] + TIER_ORDER, key="_reg_tier_filter")
    search_q = st.text_input("Search APIs", placeholder="ADMET, genomics, pathway...", key="_reg_search")
    for tier in TIER_ORDER:
        if tier_filter != "All" and tier_filter != tier:
            continue
        meta = TIER_META.get(tier, {})
        color = meta.get("color", "#94a3b8")
        tier_apis = {k: v for k, v in API_REGISTRY.items()
                     if v["tier"] == tier and
                     (not search_q or
                      search_q.lower() in v["name"].lower() or
                      search_q.lower() in v["description"].lower() or
                      search_q.lower() in v["category"].lower())}
        if not tier_apis:
            continue
        st.markdown(
            f'<div style="color:{color};font-family:JetBrains Mono,monospace;'
            f'font-size:.6rem;letter-spacing:3px;text-transform:uppercase;'
            f'margin:14px 0 6px">{meta.get("icon","")} {meta.get("label",tier)} '
            f'({len(tier_apis)} APIs)</div>',
            unsafe_allow_html=True)
        for key, entry in tier_apis.items():
            req_key = " ⚠️ Requires API Key" if entry.get("requires_key") else ""
            with st.expander(f'{entry["icon"]} {entry["name"]}{req_key}'):
                st.markdown("".join([
                    _kv("Category", entry["category"], mono=False),
                    _kv("Description", entry["description"], mono=False),
                    _kv("Base URL", entry["base_url"]),
                    _kv("Timeout", f'{entry["timeout"]}s'),
                    _kv("Requires API Key", "Yes ⚠️" if entry["requires_key"] else "No (Free)"),
                    _kv("Output Fields", " · ".join(entry.get("output_fields", []))),
                ]), unsafe_allow_html=True)
                st.markdown(f'<a href="{entry["docs"]}" target="_blank" style="color:#38bdf8;font-size:.72rem">📖 View API Documentation →</a>', unsafe_allow_html=True)
