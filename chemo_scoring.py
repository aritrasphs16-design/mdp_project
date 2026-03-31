
import math

def calculate_chemo_score(results, weights=None):
    """
    results: dict from run_comprehensive_screening (Vanguard Engine)
    weights: dict with weights for different categories
    Returns a dictionary (Package) with:
      - score: float (0-100)
      - grade: str (A, B, C, F)
      - components: dict (breakdown of scores)
    """
    if not results: return {"score": 0.0, "grade": "F", "components": {}, "weights": {}}
    
    # Extract sub-dicts if they exist (Vanguard format)
    # If it's the legacy format, fallback is handled via .get()
    props = results.get("props", results)
    rules = results.get("rules", results)
    intel = results.get("intel", results)
    
    if weights is None:
        weights = {
            "integrity": 0.20,
            "physchem": 0.25,
            "potency": 0.25, # High QED/Lead-likeness
            "safety": 0.20,
            "synthesis": 0.10
        }
    
    # 1. Integrity Score (0-1)
    # We check if it's organic, connected, and has valid valence
    # Since these are in _chemo_tests or props, we use high-level results
    i_score = 1.0
    if not props.get("organic", True): i_score -= 0.5
    if props.get("disconnected", False): i_score -= 0.3
    if props.get("unusual_valency", False): i_score -= 0.2
    i_score = max(0.0, i_score)
    
    # 2. PhysChem Score (0-1)
    # Balance MW and LogP
    mw = props.get("MW", props.get("mw", 400))
    lp = props.get("LogP", props.get("logp", 3.0))
    
    # Targeted MW ~350, LogP ~2.5
    mw_score = 1.0 - min(1.0, abs(mw - 350) / 400)
    lp_score = 1.0 - min(1.0, abs(lp - 2.5) / 5.0)
    p_score = (mw_score * 0.5) + (lp_score * 0.5)
    
    # 3. Potency/Discovery Score (0-1)
    # QED and Lipophilic Efficiency
    qed = props.get("QED", props.get("qed", 0.5))
    lipe = intel.get("Lipophilic_Efficiency", 2.0)
    lipe_score = min(1.0, max(0.0, lipe / 5.0))
    d_score = (qed * 0.6) + (lipe_score * 0.4)
    
    # 4. Safety Score (0-1)
    # PAINS, Brenk, Toxicophores
    s_score = 1.0
    if results.get("alerts", {}).get("categories", {}).get("Toxicophores", 0) > 0: s_score -= 0.4
    if results.get("alerts", {}).get("categories", {}).get("Reactive_Metabolites", 0) > 0: s_score -= 0.3
    
    # Check PAINS/Brenk from the screening results
    # safety dict in run_comprehensive_screening wasn't directly returned but repackaged
    # So we check if we can find it in the repackaged tests or just use alerts
    total_alerts = results.get("alerts", {}).get("total_hits", 0)
    s_score -= min(0.3, total_alerts * 0.05)
    s_score = max(0.0, s_score)
    
    # 5. Synthesis Score (0-1)
    # Based on SA_Score (1-10, lower is better)
    sa = props.get("SA_Score", props.get("sa_score", 5.0))
    syn_score = 1.0 - (min(10.0, max(1.0, sa)) - 1.0) / 9.0
    
    components = {
        "Integrity": round(i_score, 2),
        "PhysChem": round(p_score, 2),
        "Potency": round(d_score, 2),
        "Safety": round(s_score, 2),
        "Synthesis": round(syn_score, 2)
    }
    
    weighted_total = (
        i_score * weights.get("integrity", 0.2) +
        p_score * weights.get("physchem", 0.25) +
        d_score * weights.get("potency", 0.25) +
        s_score * weights.get("safety", 0.2) +
        syn_score * weights.get("synthesis", 0.1)
    )
    
    final_score = round(weighted_total * 100, 2)
    grade = get_grade(final_score)
    
    return {
        "score": final_score,
        "grade": grade,
        "components": components,
        "weights": weights
    }

def get_grade(score):
    if score >= 80: return "A"
    if score >= 60: return "B"
    if score >= 40: return "C"
    return "F"

def get_chemoscore_pkg(results):
    """Direct alias/wrapper for app.py and chemo_batch.py compatibility."""
    return calculate_chemo_score(results)
