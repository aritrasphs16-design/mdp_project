"""
search_engine.py
Universal search module: queries chemical_intelligence_db, master_drug_atlas,
and omnipotent_reactivity_db with fuzzy matching support.
"""

import re
from typing import Any, Dict, List, Optional

# ── Lazy imports to avoid circular deps ──────────────────────────────────────
_cid = None
_mda = None
_odb = None


def _load_dbs():
    global _cid, _mda, _odb
    if _cid is None:
        try:
            import chemical_intelligence_db as cid_mod
            _cid = cid_mod
        except Exception:
            pass
    if _mda is None:
        try:
            import master_drug_atlas as mda_mod
            _mda = mda_mod
        except Exception:
            pass
    if _odb is None:
        try:
            import omnipotent_reactivity_db as odb_mod
            _odb = odb_mod
        except Exception:
            pass


# ── Fuzzy matching helpers ────────────────────────────────────────────────────

def _fuzzy_score(query: str, target: str) -> float:
    """
    Simple fuzzy score: 1.0 = exact, 0.0 = no match.
    Uses substring containment + character overlap ratio.
    """
    q = query.lower().strip()
    t = target.lower().strip()
    if not q or not t:
        return 0.0
    if q == t:
        return 1.0
    if q in t or t in q:
        return 0.85
    # Character bigram overlap
    def bigrams(s):
        return set(s[i:i+2] for i in range(len(s) - 1))
    qb, tb = bigrams(q), bigrams(t)
    if not qb or not tb:
        return 0.0
    overlap = len(qb & tb) / max(len(qb), len(tb))
    return round(overlap, 3)


def _search_dict(query: str, data: dict, key_field: str = None,
                 threshold: float = 0.3) -> List[Dict]:
    """Search a dict {name: value} with fuzzy key matching."""
    results = []
    for name, value in data.items():
        score = _fuzzy_score(query, str(name))
        if score >= threshold:
            results.append({
                "name": name,
                "data": value,
                "score": score,
                "source": key_field or "unknown",
            })
    results.sort(key=lambda x: x["score"], reverse=True)
    return results


def search_drug_atlas(query: str, top_n: int = 10, threshold: float = 0.3) -> List[Dict]:
    """Search master_drug_atlas for matching drug names."""
    _load_dbs()
    if _mda is None:
        return []
    try:
        atlas = _mda.get_master_atlas()
        return _search_dict(query, atlas, key_field="master_drug_atlas", threshold=threshold)[:top_n]
    except Exception:
        return []


def search_chemical_db(query: str, top_n: int = 10, threshold: float = 0.3) -> List[Dict]:
    """Search chemical_intelligence_db FDA map for matching compounds."""
    _load_dbs()
    if _cid is None:
        return []
    try:
        fda_map = _cid.get_fda_map()
        return _search_dict(query, fda_map, key_field="chemical_intelligence_db",
                            threshold=threshold)[:top_n]
    except Exception:
        return []


def search_reactivity_db(query: str, top_n: int = 10, threshold: float = 0.3) -> List[Dict]:
    """Search omnipotent_reactivity_db metabolic DB."""
    _load_dbs()
    if _odb is None:
        return []
    try:
        metabolic = _odb.get_metabolic_db()
        return _search_dict(query, metabolic, key_field="reactivity_db",
                            threshold=threshold)[:top_n]
    except Exception:
        return []


def search_all(
    query: str,
    top_n: int = 10,
    threshold: float = 0.3,
    sources: Optional[List[str]] = None,
) -> Dict[str, List[Dict]]:
    """
    Search across all available databases.
    sources: optional list of ["drug_atlas", "chemical_db", "reactivity_db"]
             defaults to all.
    Returns dict keyed by source name.
    """
    all_sources = sources or ["drug_atlas", "chemical_db", "reactivity_db"]
    results = {}
    if "drug_atlas" in all_sources:
        results["drug_atlas"] = search_drug_atlas(query, top_n, threshold)
    if "chemical_db" in all_sources:
        results["chemical_db"] = search_chemical_db(query, top_n, threshold)
    if "reactivity_db" in all_sources:
        results["reactivity_db"] = search_reactivity_db(query, top_n, threshold)
    return results


def search_smiles_pattern(smarts: str, top_n: int = 10) -> List[Dict]:
    """
    Search for compounds matching a SMARTS pattern in chemical_intelligence_db.
    Requires RDKit.
    """
    _load_dbs()
    results = []
    try:
        from rdkit import Chem
        pattern = Chem.MolFromSmarts(smarts)
        if pattern is None:
            return []
        if _cid is None:
            return []
        fda_map = _cid.get_fda_map()
        for name, data in fda_map.items():
            smiles = data[0] if isinstance(data, (list, tuple)) else str(data)
            mol = Chem.MolFromSmiles(smiles)
            if mol and mol.HasSubstructMatch(pattern):
                results.append({"name": name, "smiles": smiles, "score": 1.0,
                                 "source": "chemical_intelligence_db"})
            if len(results) >= top_n:
                break
    except Exception:
        pass
    return results


def get_search_summary(results: Dict[str, List[Dict]]) -> dict:
    """Return summary stats for search_all() results."""
    total = sum(len(v) for v in results.values())
    return {
        "total_hits": total,
        "by_source": {k: len(v) for k, v in results.items()},
        "top_result": max(
            (item for items in results.values() for item in items),
            key=lambda x: x.get("score", 0),
            default=None,
        ),
    }
