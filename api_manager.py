"""
api_manager.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · API Manager — Execution Control, Caching & Failure Handling
• Orchestrates all API calls across 3 tiers
• Unified response format for all APIs
• Session-state enabled/disabled tracking
• Per-API result caching (24h TTL via st.cache_data)
• Hard timeout + graceful fallback on every call
• System works 100% offline if all APIs fail
────────────────────────────────────────────────────────────────────────────
"""

import streamlit as st
import time
import hashlib

try:
    import requests as _req
    _REQ_OK = True
except Exception:
    _req = None
    _REQ_OK = False

# ── Reliability layer (fault-tolerant wrapper, health tracking, logging) ──────
try:
    from api_reliability import (
        safe_api_call as _safe_api_call,
        safe_get as _reliability_get,
        safe_post as _reliability_post,
        get_fallback as _get_fallback,
        _update_health,
        _audit,
    )
    _RELIABILITY_OK = True
except Exception:
    _RELIABILITY_OK = False
    def _safe_api_call(fn, *a, **kw): return fn(*a, **{k: v for k, v in kw.items() if k not in ("source","fallback_data")})
    def _get_fallback(k, **kw): return {"status": "failed", "data": {}, "source": k, "error": "offline"}
    def _update_health(*a, **kw): pass
    def _audit(*a, **kw): pass

try:
    from api_registry import API_REGISTRY, TIER_META, TIER_ORDER
    _REG_OK = True
except Exception:
    _REG_OK = False
    API_REGISTRY = {}
    TIER_META = {}
    TIER_ORDER = []

_TIMEOUT_DEFAULT = 5  # seconds

# ─────────────────────────────────────────────────────────────────────────────
# UNIFIED RESPONSE FORMAT
# All API calls return this structure regardless of source
# ─────────────────────────────────────────────────────────────────────────────

def _ok(api_key: str, data: dict) -> dict:
    return {
        "status":    "ok",
        "api":       api_key,
        "data":      data,
        "error":     None,
        "timestamp": int(time.time()),
    }


def _fail(api_key: str, error: str) -> dict:
    return {
        "status":    "failed",
        "api":       api_key,
        "data":      {},
        "error":     str(error)[:200],
        "timestamp": int(time.time()),
    }


def _disabled(api_key: str) -> dict:
    return {
        "status":    "disabled",
        "api":       api_key,
        "data":      {},
        "error":     "API not enabled. Click to fetch.",
        "timestamp": 0,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE MANAGEMENT
# ─────────────────────────────────────────────────────────────────────────────

def _ss_key(api_key: str, smiles: str) -> str:
    h = hashlib.md5((api_key + smiles).encode(), usedforsecurity=False).hexdigest()[:10]
    return f"_apimgr_{api_key}_{h}"


def is_enabled(api_key: str) -> bool:
    """Check if an API is enabled in session state (user-toggled)."""
    entry = API_REGISTRY.get(api_key, {})
    default = entry.get("enabled", False)
    return st.session_state.get(f"_api_enabled_{api_key}", default)


def set_enabled(api_key: str, state: bool):
    st.session_state[f"_api_enabled_{api_key}"] = state


def get_cached_result(api_key: str, smiles: str) -> dict | None:
    """Return cached result if available, else None."""
    return st.session_state.get(_ss_key(api_key, smiles))


def store_result(api_key: str, smiles: str, result: dict):
    st.session_state[_ss_key(api_key, smiles)] = result


# ─────────────────────────────────────────────────────────────────────────────
# SAFE HTTP GET/POST WRAPPERS
# ─────────────────────────────────────────────────────────────────────────────

def _safe_get(url: str, params: dict | None = None,
              timeout: int = _TIMEOUT_DEFAULT) -> tuple:
    """Returns (response_json, error_string). Routes through reliability layer."""
    if _RELIABILITY_OK:
        import urllib.parse as _up
        full = url + ("?" + _up.urlencode(params) if params else "")
        result = _reliability_get(full, timeout=timeout, source="http")
        if result.get("status") == "success":
            return result.get("data", {}), None
        return None, result.get("error", "Request failed")
    # Fallback: direct call when reliability layer missing
    if not _REQ_OK:
        return None, "requests library not available"
    try:
        r = _req.get(url, params=params, timeout=timeout)
        r.raise_for_status()
        return r.json(), None
    except Exception as e:
        return None, str(e)[:150]


def _safe_post(url: str, json_body: dict,
               timeout: int = _TIMEOUT_DEFAULT) -> tuple:
    if _RELIABILITY_OK:
        result = _reliability_post(url, json_body, timeout=timeout, source="http")
        if result.get("status") == "success":
            return result.get("data", {}), None
        return None, result.get("error", "Request failed")
    if not _REQ_OK:
        return None, "requests library not available"
    try:
        r = _req.post(url, json=json_body, timeout=timeout)
        r.raise_for_status()
        return r.json(), None
    except Exception as e:
        return None, str(e)[:150]


# ─────────────────────────────────────────────────────────────────────────────
# INDIVIDUAL API FETCH FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

import urllib.parse as _up


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_pubchem(smiles: str) -> dict:
    enc = _up.quote(smiles.strip())
    base = "https://pubchem.ncbi.nlm.nih.gov/rest/pug"
    data, err = _safe_get(f"{base}/compound/smiles/{enc}/cids/JSON", timeout=5)
    if err or not data:
        return _fail("pubchem", err or "No CID returned")
    cids = data.get("IdentifierList", {}).get("CID", [])
    if not cids:
        return _fail("pubchem", "Compound not found in PubChem")
    cid = cids[0]
    props = "MolecularFormula,MolecularWeight,CanonicalSMILES,InChIKey,IUPACName,XLogP,TPSA,HBondDonorCount,HBondAcceptorCount,RotatableBondCount,Complexity,Charge"
    pdata, err2 = _safe_get(f"{base}/compound/cid/{cid}/property/{props}/JSON", timeout=5)
    if err2:
        return _fail("pubchem", err2)
    p = pdata.get("PropertyTable", {}).get("Properties", [{}])[0]
    syn_data, _ = _safe_get(f"{base}/compound/cid/{cid}/synonyms/JSON", timeout=4)
    syns = syn_data.get("InformationList", {}).get("Information", [{}])[0].get("Synonym", [])[:6] if syn_data else []
    return _ok("pubchem", {
        "cid": cid,
        "name": syns[0] if syns else p.get("IUPACName", "N/A"),
        "iupac_name": p.get("IUPACName", "N/A"),
        "formula": p.get("MolecularFormula", "N/A"),
        "synonyms": syns,
        "canonical_smiles": p.get("CanonicalSMILES"),
        "inchikey": p.get("InChIKey", "N/A"),
        "xlogp": p.get("XLogP"),
        "tpsa": p.get("TPSA"),
        "hbd": p.get("HBondDonorCount"),
        "hba": p.get("HBondAcceptorCount"),
        "mw": p.get("MolecularWeight"),
        "rotatable_bonds": p.get("RotatableBondCount"),
        "complexity": p.get("Complexity"),
        "charge": p.get("Charge"),
        "pubchem_url": f"https://pubchem.ncbi.nlm.nih.gov/compound/{cid}",
    })


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_chembl(smiles: str) -> dict:
    enc = _up.quote(smiles.strip())
    data, err = _safe_get(f"https://www.ebi.ac.uk/chembl/api/data/similarity/{enc}/70.json?limit=5", timeout=6)
    if err or not data:
        return _fail("chembl", err or "No response")
    mols = data.get("molecules", [])
    if not mols:
        return _fail("chembl", "No similar compounds found in ChEMBL")
    top = mols[0]
    cid = top.get("molecule_chembl_id", "")
    bio_data, _ = _safe_get(f"https://www.ebi.ac.uk/chembl/api/data/activity.json?molecule_chembl_id={cid}&limit=10", timeout=6)
    activities = []
    for a in (bio_data.get("activities", []) if bio_data else [])[:8]:
        if a.get("standard_value") and a.get("target_pref_name"):
            activities.append({
                "target": a.get("target_pref_name", "Unknown"),
                "type": a.get("standard_type", "Activity"),
                "value": a.get("standard_value"),
                "units": a.get("standard_units", ""),
                "year": a.get("document_year"),
            })
    return _ok("chembl", {
        "chembl_id": cid,
        "mol_name": top.get("pref_name") or cid,
        "similarity": top.get("similarity", 0),
        "bioactivities": activities,
        "assay_count": len(activities),
        "chembl_url": f"https://www.ebi.ac.uk/chembl/compound_report_card/{cid}/",
    })


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_pdb(query: str) -> dict:
    payload = {
        "query": {"type": "terminal", "service": "full_text",
                  "parameters": {"value": query.strip()}},
        "return_type": "entry",
        "request_options": {"paginate": {"start": 0, "rows": 5}},
    }
    data, err = _safe_post("https://search.rcsb.org/rcsbsearch/v2/query", payload, timeout=5)
    if err or not data:
        return _fail("pdb", err or "No response")
    entries = data.get("result_set", [])
    if not entries:
        return _fail("pdb", f"No PDB entries found for '{query}'")
    return _ok("pdb", {
        "entries": [{"pdb_id": e.get("identifier", ""), "score": round(e.get("score", 0), 3),
                     "url": f"https://www.rcsb.org/structure/{e.get('identifier','')}"}
                    for e in entries[:5]],
        "total_count": data.get("total_count", len(entries)),
    })


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_openfda(drug_name: str) -> dict:
    name_enc = _up.quote(drug_name.strip())
    data, err = _safe_get(f"https://api.fda.gov/drug/label.json?search=openfda.brand_name:{name_enc}&limit=1", timeout=5)
    if err or not data:
        return _fail("openfda", err or "No FDA label data")
    results = data.get("results", [{}])
    if not results:
        return _fail("openfda", "No FDA label found")
    r = results[0]
    return _ok("openfda", {
        "brand_name": r.get("openfda", {}).get("brand_name", ["N/A"])[0],
        "generic_name": r.get("openfda", {}).get("generic_name", ["N/A"])[0],
        "manufacturer": r.get("openfda", {}).get("manufacturer_name", ["N/A"])[0],
        "indications": (r.get("indications_and_usage", [""])[0])[:500],
        "warnings": (r.get("warnings", [""])[0])[:500],
        "route": r.get("openfda", {}).get("route", ["N/A"])[0],
    })


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_uniprot(gene_or_protein: str) -> dict:
    query = _up.quote(gene_or_protein.strip())
    data, err = _safe_get(
        f"https://rest.uniprot.org/uniprotkb/search?query={query}&format=json&size=3",
        timeout=5)
    if err or not data:
        return _fail("uniprot", err or "No response")
    results = data.get("results", [])
    if not results:
        return _fail("uniprot", "No UniProt entry found")
    entries = []
    for r in results[:3]:
        entries.append({
            "accession":  r.get("primaryAccession", ""),
            "name":       r.get("uniProtkbId", ""),
            "protein":    r.get("proteinDescription", {}).get("recommendedName", {}).get("fullName", {}).get("value", "N/A"),
            "organism":   r.get("organism", {}).get("scientificName", "N/A"),
            "function":   next((c["texts"][0]["value"][:300] for c in r.get("comments", [])
                                if c.get("commentType") == "FUNCTION" and c.get("texts")), "N/A"),
        })
    return _ok("uniprot", {"entries": entries})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_kegg(compound_name: str) -> dict:
    data, err = _safe_get(f"https://rest.kegg.jp/find/compound/{_up.quote(compound_name)}", timeout=5)
    if err or data is None:
        return _fail("kegg", err or "No response")
    # KEGG returns text, not JSON
    try:
        lines = data.strip().split("\n") if isinstance(data, str) else []
    except Exception:
        return _fail("kegg", "Unexpected response format")
    entries = []
    for line in lines[:5]:
        parts = line.split("\t")
        if len(parts) >= 2:
            entries.append({"id": parts[0], "name": parts[1]})
    return _ok("kegg", {"compounds": entries})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_europe_pmc(query: str) -> dict:
    data, err = _safe_get(
        "https://www.ebi.ac.uk/europepmc/webservices/rest/search",
        params={"query": query, "format": "json", "pageSize": "5", "resultType": "core"},
        timeout=6)
    if err or not data:
        return _fail("europe_pmc", err or "No response")
    results = data.get("resultList", {}).get("result", [])
    if not results:
        return _fail("europe_pmc", "No publications found")
    papers = []
    for r in results[:5]:
        papers.append({
            "pmid":     r.get("pmid", ""),
            "title":    r.get("title", "N/A"),
            "journal":  r.get("journalTitle", "N/A"),
            "year":     r.get("pubYear", ""),
            "abstract": (r.get("abstractText", "") or "")[:400],
            "doi":      r.get("doi", ""),
        })
    return _ok("europe_pmc", {"papers": papers, "hit_count": data.get("hitCount", 0)})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_clinicaltrials(query: str) -> dict:
    data, err = _safe_get(
        "https://clinicaltrials.gov/api/v2/studies",
        params={"query.term": query, "pageSize": 5, "format": "json"},
        timeout=6)
    if err or not data:
        return _fail("clinicaltrials", err or "No response")
    studies = data.get("studies", [])
    if not studies:
        return _fail("clinicaltrials", "No clinical trials found")
    trials = []
    for s in studies[:5]:
        pm = s.get("protocolSection", {})
        trials.append({
            "nct_id":   pm.get("identificationModule", {}).get("nctId", ""),
            "title":    pm.get("identificationModule", {}).get("briefTitle", "N/A"),
            "phase":    pm.get("designModule", {}).get("phases", ["N/A"])[0],
            "status":   pm.get("statusModule", {}).get("overallStatus", "N/A"),
            "condition": pm.get("conditionsModule", {}).get("conditions", ["N/A"])[0],
            "sponsor":  pm.get("sponsorCollaboratorsModule", {}).get("leadSponsor", {}).get("name", "N/A"),
        })
    return _ok("clinicaltrials", {"trials": trials, "total": data.get("totalCount", 0)})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_disgenet(gene: str) -> dict:
    data, err = _safe_get(f"https://www.disgenet.org/api/gda/gene/{_up.quote(gene)}", timeout=6)
    if err or not data:
        return _fail("disgenet", err or "No response")
    if isinstance(data, list):
        items = data[:5]
    elif isinstance(data, dict):
        items = data.get("payload", data.get("results", []))[:5]
    else:
        return _fail("disgenet", "Unexpected response")
    assocs = []
    for item in items:
        assocs.append({
            "disease":  item.get("disease_name", item.get("diseaseName", "N/A")),
            "score":    item.get("score", item.get("gda_score", "N/A")),
            "disease_id": item.get("disease_id", item.get("diseaseId", "")),
        })
    return _ok("disgenet", {"associations": assocs})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_semantic_scholar(query: str) -> dict:
    data, err = _safe_get(
        "https://api.semanticscholar.org/graph/v1/paper/search",
        params={"query": query, "limit": 5,
                "fields": "title,year,authors,citationCount,tldr,externalIds"},
        timeout=6)
    if err or not data:
        return _fail("semantic_scholar", err or "No response")
    papers = data.get("data", [])
    if not papers:
        return _fail("semantic_scholar", "No papers found")
    result = []
    for p in papers[:5]:
        result.append({
            "title":     p.get("title", "N/A"),
            "year":      p.get("year", ""),
            "authors":   ", ".join(a.get("name", "") for a in p.get("authors", [])[:3]),
            "citations": p.get("citationCount", 0),
            "tldr":      (p.get("tldr") or {}).get("text", ""),
            "doi":       p.get("externalIds", {}).get("DOI", ""),
        })
    return _ok("semantic_scholar", {"papers": result})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_chemrisk(cid: int) -> dict:
    """Fetch GHS safety data from PubChem for a given CID."""
    data, err = _safe_get(
        f"https://pubchem.ncbi.nlm.nih.gov/rest/pug_view/data/compound/{cid}/JSON?heading=GHS+Classification",
        timeout=5)
    if err or not data:
        return _fail("chemrisk", err or "No GHS data")
    try:
        sections = data.get("Record", {}).get("Section", [])
        ghs_data = {}
        for sec in sections:
            if "GHS" in sec.get("TOCHeading", ""):
                for info in sec.get("Section", []):
                    heading = info.get("TOCHeading", "")
                    values = [str(v.get("Value", {}).get("StringWithMarkup", [{}])[0].get("String", ""))
                              for v in info.get("Information", []) if v.get("Value")]
                    ghs_data[heading] = values[:5]
        return _ok("chemrisk", {"ghs_data": ghs_data, "cid": cid})
    except Exception as e:
        return _fail("chemrisk", str(e))


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_nci_cactus(name: str) -> dict:
    """NCI CACTUS — resolve drug name to SMILES."""
    enc = _up.quote(name.strip())
    data, err = _safe_get(
        f"https://cactus.nci.nih.gov/chemical/structure/{enc}/smiles",
        timeout=4)
    if err:
        return _fail("nci_cactus", err)
    smiles = data if isinstance(data, str) else str(data)
    return _ok("nci_cactus", {"smiles": smiles.strip(), "query": name})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_gtopdb(ligand_name: str) -> dict:
    data, err = _safe_get(
        f"https://www.guidetopharmacology.org/services/ligands?name={_up.quote(ligand_name)}&type=Synthetic+organic",
        timeout=5)
    if err or not data:
        return _fail("gtopdb", err or "No response")
    if not isinstance(data, list) or not data:
        return _fail("gtopdb", "No ligands found in GtoPdb")
    ligands = []
    for lig in data[:4]:
        ligands.append({
            "ligand_id":  lig.get("ligandId", ""),
            "name":       lig.get("name", "N/A"),
            "type":       lig.get("type", "N/A"),
            "approved":   lig.get("approved", False),
            "url": f"https://www.guidetopharmacology.org/GRAC/LigandDisplayForward?ligandId={lig.get('ligandId','')}",
        })
    return _ok("gtopdb", {"ligands": ligands})


@st.cache_data(ttl=86400, show_spinner=False)
def _fetch_unichem(inchikey: str) -> dict:
    data, err = _safe_get(
        f"https://www.ebi.ac.uk/unichem/api/v1/connectivity?inchikey={inchikey}",
        timeout=5)
    if err or not data:
        return _fail("unichem", err or "No response")
    references = []
    for src in (data.get("sources", []) if isinstance(data, dict) else [])[:10]:
        references.append({
            "source_name": src.get("shortName", src.get("sourceName", "N/A")),
            "compound_id": src.get("compoundId", "N/A"),
            "url": src.get("baseIdUrl", ""),
        })
    return _ok("unichem", {"cross_references": references, "inchikey": inchikey})


# ─────────────────────────────────────────────────────────────────────────────
# DISPATCH TABLE — maps api_key → fetch function
# ─────────────────────────────────────────────────────────────────────────────

_FETCH_DISPATCH = {
    "pubchem":          lambda smiles, query, **kw: _fetch_pubchem(smiles),
    "chembl":           lambda smiles, query, **kw: _fetch_chembl(smiles),
    "pdb":              lambda smiles, query, **kw: _fetch_pdb(query or "drug target"),
    "openfda":          lambda smiles, query, **kw: _fetch_openfda(query or smiles[:20]),
    "uniprot":          lambda smiles, query, **kw: _fetch_uniprot(query or ""),
    "kegg":             lambda smiles, query, **kw: _fetch_kegg(query or ""),
    "europe_pmc":       lambda smiles, query, **kw: _fetch_europe_pmc(query or smiles[:30]),
    "clinicaltrials":   lambda smiles, query, **kw: _fetch_clinicaltrials(query or ""),
    "disgenet":         lambda smiles, query, **kw: _fetch_disgenet(query or ""),
    "semantic_scholar": lambda smiles, query, **kw: _fetch_semantic_scholar(query or ""),
    "chemrisk":         lambda smiles, query, **kw: _fetch_chemrisk(kw.get("cid", 0)),
    "nci_cactus":       lambda smiles, query, **kw: _fetch_nci_cactus(query or smiles[:20]),
    "gtopdb":           lambda smiles, query, **kw: _fetch_gtopdb(query or ""),
    "unichem":          lambda smiles, query, **kw: _fetch_unichem(kw.get("inchikey", "")),
}


def fetch_api(api_key: str, smiles: str = "", query: str = "", **kwargs) -> dict:
    """
    Main dispatch function. Routes through reliability layer.
    Checks session cache → fetches → falls back on failure.
    Returns unified response dict. Never crashes.
    """
    cache_key_str = smiles + query + str(sorted(kwargs.items()))
    cached = get_cached_result(api_key, cache_key_str)
    if cached and cached.get("status") in ("ok", "success"):
        return cached
    fn = _FETCH_DISPATCH.get(api_key)
    if fn is None:
        result = _fail(api_key, "No fetch function registered for this API")
    else:
        try:
            if _RELIABILITY_OK:
                raw = _safe_api_call(
                    fn, smiles=smiles, query=query,
                    source=api_key, **kwargs)
                result = raw
            else:
                result = fn(smiles=smiles, query=query, **kwargs)
        except Exception as e:
            result = _fail(api_key, str(e))
    # If result still failed → use fallback
    if result.get("status") == "failed" and _RELIABILITY_OK:
        result = _get_fallback(api_key, smiles=smiles)
    store_result(api_key, cache_key_str, result)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# RESULT DISPLAY HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def status_badge(result: dict) -> str:
    s = result.get("status", "unknown")
    if s == "ok":      return "🟢 Data Retrieved"
    if s == "failed":  return f"🔴 Failed: {result.get('error','')}"
    if s == "disabled": return "⚪ Not Fetched"
    return "⚪ Unknown"


def render_api_status_row(api_key: str, result: dict):
    """Render a compact one-line status row for an API result."""
    entry = API_REGISTRY.get(api_key, {})
    icon  = entry.get("icon", "⬡")
    name  = entry.get("name", api_key)
    badge = status_badge(result)
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:10px;padding:4px 0;'
        f'border-bottom:1px solid rgba(255,255,255,.04)">'
        f'<span style="font-size:.85rem">{icon}</span>'
        f'<span style="font-size:.75rem;color:#c8deff;flex:1">{name}</span>'
        f'<span style="font-size:.68rem;color:#64748b">{badge}</span></div>',
        unsafe_allow_html=True,
    )
