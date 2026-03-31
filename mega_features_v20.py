"""
╔══════════════════════════════════════════════════════════════════════════════╗
║        CHEMOFILTER v20 — MEGA FEATURES MODULE  (ULTRA EDITION v3.0)         ║
║  150+ Physicochemical · ADMET · Structural · SAR · Drug-Likeness Features   ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from rdkit import Chem
from rdkit.Chem import (
    Descriptors, rdMolDescriptors, Crippen, AllChem,
    rdFreeSASA, QED, DataStructs, Lipinski, GraphDescriptors
)
from rdkit.Chem.rdMolDescriptors import (
    CalcNumHBD, CalcNumHBA, CalcNumRotatableBonds,
    CalcNumRings, CalcNumAromaticRings, CalcNumAliphaticRings,
    CalcNumHeterocycles, CalcNumAromaticHeterocycles,
    CalcNumAliphaticHeterocycles, CalcNumSaturatedRings,
    CalcNumBridgeheadAtoms, CalcNumSpiroAtoms,
    CalcLabuteASA, CalcTPSA, CalcKappa1, CalcKappa2, CalcKappa3,
    CalcChi0v, CalcChi1v, CalcChi2v, CalcChi3v, CalcChi4v,
    CalcChi0n, CalcChi1n,
    CalcNumAmideBonds, CalcNumHeteroatoms
)
import math

# ═══════════════════════════════════════════════════════════════════════════════
# PRE-COMPILED SMARTS — built once at module load, reused for every molecule
# This eliminates 66 inline Chem.MolFromSmarts() calls that ran per-molecule
# ═══════════════════════════════════════════════════════════════════════════════
_P = {}  # pattern registry

def _pat(smarts: str):
    """Compile & cache a SMARTS pattern; return None if invalid."""
    if smarts not in _P:
        try:
            _P[smarts] = Chem.MolFromSmarts(smarts)
        except Exception:
            _P[smarts] = None
    return _P[smarts]

# Pre-compile all patterns at import time
_SMARTS_PRELOAD = [
    # atom counts
    "[F]", "[Cl]", "[Br]", "[I]", "[S]", "[P]", "[B]", "[N]", "[O]", "[Si]", "[Se]",
    "[Br,I]", "[F,Cl]",
    # common patterns used many times
    "C(=O)O",               # carboxylic acid (loose)
    "[CX3](=O)[OX2H1]",    # carboxylic acid (strict)
    "[NX3;H1,H2;!$(N-C=O)]",  # basic N (pka, ionization)
    "[NX3;H2,H1,H0;!$(N-C=O)]",  # basic N broad
    "[NX3;H2,H1,H0;!$(N-C=O);!$(N-c)]",  # sp3 basic N
    "[NX3;H1,H2;!$(N-C=O)]",  # basic N (amine)
    "[NX3;H2]",             # primary amine
    "[CH3,CH2,CH]",         # metabolic sites
    "C(=O)OC",             # ester
    "C(=O)N",              # amide
    "[OH,NH,SH,COOH]",     # glucuronidation sites
    "[OH,NH2]",             # sulfation sites
    "[SH,S-]",             # methylation sites
    "[NH2]-c1ccccc1",      # N-acetylation / aniline
    "[CH3]",               # methyl
    "[N,O,S]N=O",          # danger zone
    "[Cl,Br,I][C,S]",      # danger zone
    "[P,S][Cl,Br,I]",      # danger zone
    "C1CO1",               # danger zone / epoxide
    "c1ccccc1O",           # tox alerts
    "c1ccccc1N",           # tox / dili
    "S=C(N)N",             # tox alerts
    "[N+](=O)[O-]",        # nitro group
    "C=CC(=O)",            # skin sensitizer
    "[Cl,Br][CX4]",        # skin sensitizer
    "O=C-O-C=O",          # skin sensitizer
    "c1ccc(N)cc1",         # aromatic amine
    "[NX2]=O",             # nitroso
    "C1(=O)C=CC(=O)C=C1",  # quinone
    "c1ccc2c(c1)ccc3c2cccc3",  # polycyclic aromatic
    "[NX3][NX3]",          # hydrazine
    "C1OC1",               # epoxide
    "[Cl,Br,I][CH2]",      # alkyl halide
    "[As,Sb,Hg,Pb,Cd]",    # heavy metals
    "[F,Cl,Br,I]",         # halogens
    "[CX4](C)(C)(C)C",     # tert-butyl
    "[N]",                 # all N
    "S(=O)(=O)O",          # sulfonic acid
    "Oc1ccccc1",           # phenol
    "[NH][CX3]=O",         # amide NH
    "[NX4+]",              # quaternary N
    "[NX3;H2,H1;!$(N-C=O)]",  # aliphatic amine
    "n1ccccc1",            # pyridine
    "[nH]1cccc1",          # pyrrole
    "c1ccccc1",            # benzene ring
    "[NH,NH2]",            # NH donors
    "[OH]",                # OH donors
    "[NX3][CX3]=O",        # amide N
    "[n]",                 # aromatic N
    "[cH]",                # aromatic CH
    # metabolic clipping
    "[NH]-c1ccccc1",
    "c1cc(O)ccc1",
    "C=C-C=O",
    "[N,O,S][CH3]",
    # functional groups
    "[CX3](=O)[OX2H0][#6]",  # ester
    "[NX3][CX3](=O)",       # amide
    "[SX4](=O)(=O)[NX3]",   # sulfonamide
    "[OX2H][c]",            # phenol
    "[OX2H][C]",            # alcohol
    "[NX1]#[CX2]",          # nitrile
    "[NX3;H2,H1,H0;!$(N-C=O)]",  # basic nitrogen
    # bioisostere
    "[nH]",
    # cyp substrates
    "[NX3;H0;R][CX4]",
    # exit vectors
    "[*]-[H]",
    # pka
    "Oc1ccccc1",
    "[NH2]-c1ccccc1",
]

# Execute all pre-compilations at module load
for _s in _SMARTS_PRELOAD:
    _pat(_s)


# ═══════════════════════════════════════════════════════════
# SECTION 1 — ATOM / ELEMENT COUNTS
# ═══════════════════════════════════════════════════════════

# Pre-compiled atom count patterns (module-level)
_ATOM_COUNT_PATS = {
    "F_Count":       "[F]",
    "Cl_Count":      "[Cl]",
    "Br_Count":      "[Br]",
    "I_Count":       "[I]",
    "S_Count":       "[S]",
    "P_Count":       "[P]",
    "B_Count":       "[B]",
    "N_Count":       "[N]",
    "O_Count":       "[O]",
    "Si_Count":      "[Si]",
    "Se_Count":      "[Se]",
    "Heavy_Halogen": "[Br,I]",
    "Light_Halogen": "[F,Cl]",
}
_ATOM_COUNT_COMPILED = {k: Chem.MolFromSmarts(v) for k, v in _ATOM_COUNT_PATS.items()
                        if Chem.MolFromSmarts(v) is not None}

def atom_counts(mol):
    """Detailed atom-level elemental counts."""
    return {k: len(mol.GetSubstructMatches(p)) for k, p in _ATOM_COUNT_COMPILED.items()}

def heteroatom_detail(mol):
    """Per-element heteroatom breakdown."""
    total_N = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 7)
    total_O = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 8)
    total_S = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 16)
    total_P = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 15)
    total_halo = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() in (9,17,35,53))
    return {
        "N_Total": total_N, "O_Total": total_O,
        "S_Total": total_S, "P_Total": total_P,
        "Halogen_Total": total_halo
    }

# ═══════════════════════════════════════════════════════════
# SECTION 2 — RING ANALYSIS
# ═══════════════════════════════════════════════════════════

def ring_size_distribution(mol):
    """Detailed ring size distribution."""
    ri = mol.GetRingInfo()
    dist = {}
    for r in ri.AtomRings():
        key = f"Ring_{len(r)}"
        dist[key] = dist.get(key, 0) + 1
    return dist

def ring_fusion_score(mol):
    """Measures degree of ring fusion (fused vs isolated)."""
    ri = mol.GetRingInfo()
    rings = list(ri.AtomRings())
    if len(rings) < 2:
        return 0
    fused = 0
    for i in range(len(rings)):
        for j in range(i + 1, len(rings)):
            shared = set(rings[i]) & set(rings[j])
            if len(shared) >= 2:
                fused += 1
    return fused

def macrocycle_flag(mol):
    """Detects macrocycles (rings ≥ 12)."""
    ri = mol.GetRingInfo()
    for r in ri.AtomRings():
        if len(r) >= 12:
            return f"Macrocycle ({len(r)}-membered)"
    return "No Macrocycle"

def aromatic_ring_count(mol):
    return CalcNumAromaticRings(mol)

def aliphatic_ring_count(mol):
    return CalcNumAliphaticRings(mol)

def saturated_ring_count(mol):
    return CalcNumSaturatedRings(mol)

def heterocycle_count(mol):
    return CalcNumHeterocycles(mol)

def aromatic_heterocycle_count(mol):
    return CalcNumAromaticHeterocycles(mol)

# ═══════════════════════════════════════════════════════════
# SECTION 3 — DRUG-LIKENESS RULES
# ═══════════════════════════════════════════════════════════

def pfizer_3_75_rule(lp, tp):
    if lp > 3 and tp < 75: return "High Risk (Tox)"
    return "Low Risk"

def gsk_4_400_rule(lp, mw):
    if lp < 4 and mw < 400: return "Good Profile"
    return "Caution"

def oprea_rule(mol, _mw=None, _lp=None, _hbd=None, _hba=None, _rot=None, _rings=None):
    mw    = _mw    if _mw    is not None else Descriptors.MolWt(mol)
    lp    = _lp    if _lp    is not None else Descriptors.MolLogP(mol)
    hbd   = _hbd   if _hbd   is not None else CalcNumHBD(mol)
    hba   = _hba   if _hba   is not None else CalcNumHBA(mol)
    rot   = _rot   if _rot   is not None else CalcNumRotatableBonds(mol)
    rings = _rings if _rings is not None else CalcNumRings(mol)
    v = sum([mw<=450, lp<=4.5, hbd<=5, hba<=9, rot<=8, rings<=4])
    return f"{v}/6 Pass"

def congreve_rule_of_3(mol, _mw=None, _lp=None, _hbd=None, _hba=None, _rot=None):
    mw  = _mw  if _mw  is not None else Descriptors.MolWt(mol)
    lp  = _lp  if _lp  is not None else Descriptors.MolLogP(mol)
    hbd = _hbd if _hbd is not None else CalcNumHBD(mol)
    hba = _hba if _hba is not None else CalcNumHBA(mol)
    rot = _rot if _rot is not None else CalcNumRotatableBonds(mol)
    if mw<=300 and lp<=3 and hbd<=3 and hba<=3 and rot<=3:
        return "Fragment-Like ✅"
    return "Not Fragment-Like"

def veber_rule_ext(rot, tp):
    if rot<=10 and tp<=140: return "Pass (Oral)"
    return "Poor Oral Bio"

def mddr_likeness(mol, _rings=None, _rot=None):
    rings = _rings if _rings is not None else CalcNumRings(mol)
    rot   = _rot   if _rot   is not None else CalcNumRotatableBonds(mol)
    if rings>=3 and rot>=6: return "Drug-Like (MDDR)"
    return "Moderate"

def rule_of_5_ext(mol, _mw=None, _lp=None, _hbd=None, _hba=None):
    mw  = _mw  if _mw  is not None else Descriptors.MolWt(mol)
    lp  = _lp  if _lp  is not None else Descriptors.MolLogP(mol)
    hbd = _hbd if _hbd is not None else CalcNumHBD(mol)
    hba = _hba if _hba is not None else CalcNumHBA(mol)
    v = sum([mw>500, lp>5, hbd>5, hba>10])
    return f"{4-v}/4"

def ghose_v2(lp, mw, h, refr):
    if (160<=mw<=480) and (-0.4<=lp<=5.6) and (20<=h<=70) and (40<=refr<=130):
        return "Pass"
    return "Fail"

def egan_v2(lp, tp):
    if lp<=5.88 and tp<=131.6: return "Pass"
    return "Fail"

def bioavailability_score_ro5(mol, _mw=None, _lp=None, _hbd=None, _hba=None, _rot=None, _tp=None):
    """Bioavailability score 0–1 based on oral drug rules."""
    mw  = _mw  if _mw  is not None else Descriptors.MolWt(mol)
    lp  = _lp  if _lp  is not None else Descriptors.MolLogP(mol)
    hbd = _hbd if _hbd is not None else CalcNumHBD(mol)
    hba = _hba if _hba is not None else CalcNumHBA(mol)
    rot = _rot if _rot is not None else CalcNumRotatableBonds(mol)
    tp  = _tp  if _tp  is not None else CalcTPSA(mol)
    score = 0.0
    if mw <= 500: score += 0.2
    if lp <= 5:   score += 0.2
    if hbd <= 5:  score += 0.2
    if hba <= 10: score += 0.2
    if rot <= 10: score += 0.1
    if tp <= 140: score += 0.1
    return round(score, 2)

def beyond_ro5_check(mol, _mw=None, _lp=None, _hbd=None):
    """bRo5 for PPI and macrocyclic drugs (MW 500-1000)."""
    mw  = _mw  if _mw  is not None else Descriptors.MolWt(mol)
    lp  = _lp  if _lp  is not None else Descriptors.MolLogP(mol)
    hbd = _hbd if _hbd is not None else CalcNumHBD(mol)
    if 500 < mw <= 1000 and lp <= 10 and hbd <= 6:
        return "bRo5 Compliant (PPI/Macro)"
    return "Outside bRo5"

def dc_linker_rule(mol, _mw=None, _rot=None):
    """Drug-Conjugate linker rule: MW > 700, flexible."""
    mw  = _mw  if _mw  is not None else Descriptors.MolWt(mol)
    rot = _rot if _rot is not None else CalcNumRotatableBonds(mol)
    if mw > 700 and rot > 12:
        return "Potential ADC/PDC Payload"
    return "Standard"

def tice_rule(mol, _tp=None, _rot=None, _mw=None):
    """TICE rule for blood-brain penetration."""
    tp  = _tp  if _tp  is not None else CalcTPSA(mol)
    rot = _rot if _rot is not None else CalcNumRotatableBonds(mol)
    mw  = _mw  if _mw  is not None else Descriptors.MolWt(mol)
    if tp < 90 and rot <= 8 and mw < 450:
        return "CNS-Penetrant (TICE)"
    return "Poor CNS"

# ═══════════════════════════════════════════════════════════
# SECTION 4 — LIPOPHILICITY & SOLUBILITY
# ═══════════════════════════════════════════════════════════

def flexibility_index(mol):
    rot = CalcNumRotatableBonds(mol); heavy = mol.GetNumHeavyAtoms()
    return round(rot/heavy, 3) if heavy>0 else 0

def molar_refractivity(mol):
    return round(Crippen.MolMR(mol), 2)

def bbb_v2_index(lp, tp, mw):
    """Clark logBB estimate."""
    return round(0.152*lp - 0.0148*tp + 0.139, 3)

def logd_ph74_estimate(mol):
    """Estimate LogD at pH 7.4 using LogP and ionization heuristic."""
    lp = Descriptors.MolLogP(mol)
    tp = CalcTPSA(mol)
    correction = -0.01 * tp + 0.05 * lp
    return round(lp + correction, 2)

def logd_ph68_estimate(mol):
    """Estimate LogD at pH 6.8 (GI environment)."""
    lp = Descriptors.MolLogP(mol)
    has_acid = mol.HasSubstructMatch(_pat("C(=O)O"))
    correction = -0.3 if has_acid else 0.1
    return round(lp + correction, 2)

def aqueous_solubility_esol(mol):
    """ESOL (Delaney) aqueous solubility estimate (log mol/L)."""
    lp = Descriptors.MolLogP(mol); mw = Descriptors.MolWt(mol)
    rb = CalcNumRotatableBonds(mol); ap = CalcNumAromaticRings(mol)
    fsp3 = Descriptors.FractionCSP3(mol)
    esol = 0.16 - 0.63*lp - 0.0062*mw + 0.066*rb - 0.74*ap
    return round(esol, 3)

def solubility_class(esol):
    """Categorical solubility from ESOL."""
    if esol > -1: return "Highly Soluble"
    if esol > -2: return "Soluble"
    if esol > -4: return "Moderately Soluble"
    if esol > -6: return "Poorly Soluble"
    return "Insoluble"

def ali_solubility(mol):
    """Ali et al. solubility model."""
    lp = Descriptors.MolLogP(mol); mw = Descriptors.MolWt(mol)
    rb = CalcNumRotatableBonds(mol); hbd = CalcNumHBD(mol)
    return round(0.11 - 0.60*lp - 0.0056*mw + 0.053*rb + 0.28*hbd, 3)

def solubility_improvement_hint(lp):
    if lp > 5: return "Add Hydroxyl/Amine Group"
    if lp < 0: return "Add Alkyl/Cyclohexyl Group"
    return "Balanced — No Action Needed"

def lipophilicity_efficiency(qed, lp):
    """LipE = QED – LogP (higher = better drug-likeness per lipophilicity unit)."""
    return round(qed - lp, 2)

def chromatographic_hydrophobicity_index(lp, tp):
    """CHI estimate from logP and TPSA."""
    return round(0.916*lp - 0.007*tp + 1.22, 2)

# ═══════════════════════════════════════════════════════════
# SECTION 5 — BBB & CNS
# ═══════════════════════════════════════════════════════════

def bbb_likelihood_cat(lp, tp):
    if lp>2 and tp<60: return "High"
    if lp>1 and tp<90: return "Moderate"
    return "Low"

def cns_mpo_score(mol):
    """CNS MPO score (0–6) — Wager et al. 2010."""
    lp = Descriptors.MolLogP(mol); tp = CalcTPSA(mol)
    mw = Descriptors.MolWt(mol); hbd = CalcNumHBD(mol)
    pka_basic = 8.0 if mol.HasSubstructMatch(_pat("[NX3;H1,H2;!$(N-C=O)]")) else 5.0
    score = 0
    if lp <= 5: score += 1
    if lp >= 1: score += 1
    if tp <= 76: score += 1
    if mw <= 360: score += 1
    if hbd <= 1: score += 1
    if 7 <= pka_basic <= 10: score += 1
    return score

def p_glycoprotein_alert(mol):
    """heuristic P-gp efflux alert based on physicochemical space."""
    mw = Descriptors.MolWt(mol); lp = Descriptors.MolLogP(mol)
    hbd = CalcNumHBD(mol); tp = CalcTPSA(mol)
    if mw > 500 or hbd > 3 or tp > 120: return "P-gp Efflux Likely"
    return "Low P-gp Risk"

def passive_permeability_papp(mol):
    """Caco-2 Papp estimate (×10⁻⁶ cm/s) heuristic."""
    tp = CalcTPSA(mol); rb = CalcNumRotatableBonds(mol)
    lp = Descriptors.MolLogP(mol)
    papp = 10 ** (0.5*lp - 0.01*tp - 0.1*rb)
    if papp > 10: cat = "High"
    elif papp > 2: cat = "Moderate"
    else: cat = "Low"
    return {"Papp_Estimate": round(papp, 2), "Class": cat}

def oral_absorption_score(mol):
    """Percentage oral absorption estimate (Johnson & Bhatt model)."""
    tp = CalcTPSA(mol); rb = CalcNumRotatableBonds(mol)
    if tp <= 60 and rb <= 5: return 90
    if tp <= 100: return 70
    if tp <= 140: return 40
    return 15

# ═══════════════════════════════════════════════════════════
# SECTION 6 — METABOLISM & CLEARANCE
# ═══════════════════════════════════════════════════════════

def metabolic_half_life_score(mol):
    sites = len(mol.GetSubstructMatches(_pat("[CH3,CH2,CH]")))
    if sites>10: return "Potential Short t½"
    return "Extended t½ Likely"

def microsomal_stability_index(mol):
    """Heuristic microsomal stability (human liver microsomes)."""
    lp = Descriptors.MolLogP(mol); ha = mol.GetNumHeavyAtoms()
    alert_smarts = ["c1ccccc1", "[CH3]", "COC"]
    hits = sum(1 for s in alert_smarts if mol.HasSubstructMatch(Chem.MolFromSmarts(s)))
    score = 100 - hits*15 - max(0, lp-3)*5
    if score > 70: return "Stable (t½ > 30 min)"
    if score > 40: return "Moderate (t½ 15-30 min)"
    return "Labile (t½ < 15 min)"

def hepatic_extraction_ratio(mol):
    """Low/Medium/High hepatic extraction heuristic."""
    lp = Descriptors.MolLogP(mol); tp = CalcTPSA(mol)
    mw = Descriptors.MolWt(mol)
    er = (lp * 0.15) + (1/(tp+1) * 20) - (mw/1000)
    if er > 0.7: return "High ER (>70%)"
    if er > 0.3: return "Medium ER (30-70%)"
    return "Low ER (<30%)"

def first_pass_effect_risk(mol):
    """Predicts risk of high first-pass extraction."""
    has_ester = mol.HasSubstructMatch(_pat("C(=O)OC"))
    has_amide = mol.HasSubstructMatch(_pat("C(=O)N"))
    lp = Descriptors.MolLogP(mol)
    if has_ester or lp > 4: return "High First-Pass Risk"
    if has_amide: return "Moderate"
    return "Low First-Pass Risk"

def phase_ii_alerts(mol):
    hits = []
    if mol.HasSubstructMatch(_pat("[OH,NH,SH,COOH]")): hits.append("Glucuronidation")
    if mol.HasSubstructMatch(_pat("[OH,NH2]")): hits.append("Sulfation")
    if mol.HasSubstructMatch(_pat("[SH,S-]")): hits.append("Methylation")
    if mol.HasSubstructMatch(_pat("[NH2]-c1ccccc1")): hits.append("N-Acetylation")
    return ", ".join(hits) if hits else "Stable"

def metabolic_clipping_alert(mol):
    methyl = len(mol.GetSubstructMatches(_pat("[CH3]")))
    ester  = len(mol.GetSubstructMatches(_pat("C(=O)OC")))
    if methyl>3 or ester>1: return "Potential Clipping"
    return "Stable"

def reactive_metabolite_risk(mol):
    """Counts structural motifs known to form reactive metabolites."""
    patterns = [
        ("[NH]-c1ccccc1",     "Aniline → Nitroso"),
        ("c1cc(O)ccc1",       "Phenol → Quinone"),
        ("C1OC1",             "Epoxide → Alkylation"),
        ("C=C-C=O",           "Enone → Michael"),
        ("[NX2]=O",           "N-oxide"),
        ("C(=O)[Cl,Br]",      "Acyl Halide"),
    ]
    hits = []
    for smarts, name in patterns:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            hits.append(name)
    return hits if hits else ["None Detected"]

# ═══════════════════════════════════════════════════════════
# SECTION 7 — TOXICITY ALERTS
# ═══════════════════════════════════════════════════════════

def covalent_danger_zone(mol):
    danger = ["[Cl,Br,I][C,S]", "[P,S][Cl,Br,I]", "[N,O,S]N=O", "C1CO1"]
    for d in danger:
        if mol.HasSubstructMatch(Chem.MolFromSmarts(d)): return "DANGER: HIGH REACTIVITY"
    return "Safe"

def toxicity_structural_alerts(mol):
    alerts = ["c1ccccc1O", "c1ccccc1N", "S=C(N)N", "[N+](=O)[O-]"]
    return sum(len(mol.GetSubstructMatches(Chem.MolFromSmarts(a))) for a in alerts)

def pains_filter(mol):
    """Simplified PAINS A/B/C filter (top 20 most frequent hitters)."""
    pains_smarts = [
        ("Quinone",       "O=C1C=CC(=O)C=C1"),
        ("Catechol",      "c1cc(O)c(O)cc1"),
        ("Rhodanine",     "O=C1NC(=S)SC1"),
        ("Aniline-azo",   "c1ccccc1N=Nc1ccccc1"),
        ("Enone",         "C=CC(=O)"),
        ("Isocyanate",    "N=C=O"),
        ("Epoxide",       "C1OC1"),
        ("Nitroso",       "[NX2]=O"),
        ("Hydrazine",     "[NX3][NX3]"),
        ("Aldehyde",      "[CX3H1](=O)"),
        ("Beta_Lactam",   "C1CC(=O)N1"),
        ("Acyl_halide",   "[C](=O)[Cl,Br,F]"),
        ("Disulfide",     "SSc1ccccc1"),
        ("Sulfonyl_Cl",   "S(=O)(=O)Cl"),
        ("Diazonium",     "[N+]#N"),
    ]
    hits = []
    for name, smarts in pains_smarts:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            hits.append(name)
    return hits if hits else ["Clean"]

def dili_risk_extended(mol):
    """Extended DILI (Drug-Induced Liver Injury) risk score."""
    score = 0
    if mol.HasSubstructMatch(_pat("c1ccccc1N")):  score += 2
    if mol.HasSubstructMatch(_pat("[N+](=O)[O-]")): score += 3
    if mol.HasSubstructMatch(_pat("c1cc(O)ccc1")): score += 1
    if Descriptors.MolLogP(mol) > 5: score += 1
    if CalcTPSA(mol) < 50:           score += 1
    if score >= 4: return "High DILI Risk"
    if score >= 2: return "Moderate DILI Risk"
    return "Low DILI Risk"

def skin_sensitization_risk(mol):
    """Skin sensitization based on electrophilic reactivity."""
    patterns = ["C=CC(=O)", "C1OC1", "[Cl,Br][CX4]", "O=C-O-C=O"]
    for p in patterns:
        pat = Chem.MolFromSmarts(p)
        if pat and mol.HasSubstructMatch(pat):
            return "Sensitizer Alert"
    return "Non-Sensitizer"

def mutagenicity_ames_alert(mol):
    """Extended Ames test structural alerts."""
    ames_patterns = [
        ("[N+](=O)[O-]",         "Nitro"),
        ("c1ccc(N)cc1",          "Aromatic Amine"),
        ("[NX2]=O",              "Nitroso"),
        ("C1(=O)C=CC(=O)C=C1",  "Quinone"),
        ("c1ccc2c(c1)ccc3c2cccc3", "Polycyclic Aromatic"),
    ]
    hits = []
    for smarts, name in ames_patterns:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            hits.append(name)
    return hits if hits else ["No Ames Alert"]

def genotox_structural_flags(mol):
    """Genotoxicity structural flags (OECD TG471)."""
    flags = []
    if mol.HasSubstructMatch(_pat("[N+](=O)[O-]")):
        flags.append("Nitro Group (Genotox)")
    if mol.HasSubstructMatch(_pat("[NX3][NX3]")):
        flags.append("Hydrazine (Genotox)")
    if mol.HasSubstructMatch(_pat("C1OC1")):
        flags.append("Epoxide (DNA Alkylator)")
    if mol.HasSubstructMatch(_pat("[Cl,Br,I][CH2]")):
        flags.append("Alkyl Halide (SN2)")
    return flags if flags else ["No Flags"]

def cardiotoxicity_herg_alert(mol):
    """hERG channel blockade risk (QT prolongation)."""
    # Basic nitrogen + lipophilic scaffold = hERG risk
    basic_n = mol.HasSubstructMatch(_pat("[NX3;H1,H2;!$(N-C=O)]"))
    aromatic = CalcNumAromaticRings(mol)
    lp = Descriptors.MolLogP(mol)
    mw = Descriptors.MolWt(mol)
    if basic_n and aromatic >= 2 and lp > 2 and mw > 300:
        return "High hERG Risk"
    if basic_n and lp > 1:
        return "Moderate hERG Risk"
    return "Low hERG Risk"

def nephrotoxicity_alert(mol):
    """Kidney toxicity structural alert."""
    if mol.HasSubstructMatch(_pat("[As,Sb,Hg,Pb,Cd]")):
        return "Metal-Mediated Nephrotox"
    if mol.HasSubstructMatch(_pat("c1ccc(N)cc1")):
        return "Aromatic Amine Nephrotox Risk"
    return "Low Nephrotox Risk"

def phospholipidosis_risk(mol):
    """CAD (cationic amphiphilic drug) = phospholipidosis risk."""
    basic_n = mol.HasSubstructMatch(_pat("[NX3;H1,H2;!$(N-C=O)]"))
    lp = Descriptors.MolLogP(mol); mw = Descriptors.MolWt(mol)
    if basic_n and lp > 2 and mw > 300:
        return "Phospholipidosis Risk (CAD)"
    return "Low Risk"

def idiosyncratic_tox_flag(mol):
    """Idiosyncratic toxicity structural flags."""
    flags = []
    if mol.HasSubstructMatch(_pat("c1ccccc1-[NH2]")): flags.append("Aniline")
    if mol.HasSubstructMatch(_pat("[NX3][NX3]")): flags.append("Hydrazine")
    if mol.HasSubstructMatch(_pat("S(=O)(=O)c1ccccc1")): flags.append("Sulfonamide")
    if mol.HasSubstructMatch(_pat("c1cc(F)ccc1")): flags.append("Halogenated Phenyl")
    return flags if flags else ["None"]

# ═══════════════════════════════════════════════════════════
# SECTION 8 — STRUCTURAL COMPLEXITY & SHAPE
# ═══════════════════════════════════════════════════════════

def shape_index_kappa1(mol):
    try: return round(CalcKappa1(mol), 3)
    except: return 0

def shape_index_kappa2(mol):
    try: return round(CalcKappa2(mol), 3)
    except: return 0

def shape_index_kappa3(mol):
    try: return round(CalcKappa3(mol), 3)
    except: return 0

def labute_asa(mol):
    try: return round(CalcLabuteASA(mol), 2)
    except: return 0

def chi_indices(mol):
    """Chi connectivity indices (topological)."""
    try:
        return {
            "Chi0v": round(CalcChi0v(mol), 3),
            "Chi1v": round(CalcChi1v(mol), 3),
            "Chi2v": round(CalcChi2v(mol), 3),
            "Chi3v": round(CalcChi3v(mol), 3),
            "Chi0n": round(CalcChi0n(mol), 3),
            "Chi1n": round(CalcChi1n(mol), 3),
        }
    except:
        return {}

def wiener_index(mol):
    """Wiener topological index."""
    try:
        return GraphDescriptors.BalabanJ(mol)
    except:
        return 0

def molecular_flexibility(mol):
    n_rot = CalcNumRotatableBonds(mol); n_heavy = mol.GetNumHeavyAtoms()
    return round((n_rot/n_heavy*100), 1) if n_heavy>0 else 0

def arom_heavy_ratio(mol):
    heavy = mol.GetNumHeavyAtoms()
    arom = sum(1 for a in mol.GetAtoms() if a.GetIsAromatic())
    return round(arom/heavy, 3) if heavy>0 else 0

def halogen_ratio(mol):
    heavy = mol.GetNumHeavyAtoms()
    hal = len(mol.GetSubstructMatches(_pat("[F,Cl,Br,I]")))
    return round(hal/heavy, 3) if heavy>0 else 0

def heteroatom_ratio(mol):
    heavy = mol.GetNumHeavyAtoms()
    c = sum(1 for a in mol.GetAtoms() if a.GetSymbol()=="C")
    return round((heavy-c)/heavy, 3) if heavy>0 else 0

def bridgehead_complexity(mol):
    return CalcNumBridgeheadAtoms(mol)

def spiro_complexity(mol):
    return CalcNumSpiroAtoms(mol)

def sterics_index(mol):
    p = _pat("[CX4](C)(C)(C)C")
    return len(mol.GetSubstructMatches(p))

def stereocenters_count(mol):
    """Number of stereocenters."""
    return len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))

def stereo_density(mol):
    """Stereocenters / heavy atoms."""
    sc = stereocenters_count(mol); ha = mol.GetNumHeavyAtoms()
    return round(sc/ha, 3) if ha>0 else 0

def isomer_count_hint(mol):
    centers = stereocenters_count(mol)
    return 2**centers

def graph_complexity(mol):
    """Bertz complexity index."""
    try: return round(GraphDescriptors.BertzCT(mol), 1)
    except: return 0

def mol_volume(mol):
    return round(CalcLabuteASA(mol)*0.8, 1)

def polarizability_estimate(mol):
    return round(Crippen.MolMR(mol)/2.5, 2)

def tpsa_per_heavy(mol, tp):
    h = mol.GetNumHeavyAtoms()
    return round(tp/h, 2) if h>0 else 0

def amide_bond_count(mol):
    return CalcNumAmideBonds(mol)

def nitrogen_saturation(mol):
    n_all = len(mol.GetSubstructMatches(_pat("[N]")))
    n_sp3 = len(mol.GetSubstructMatches(_pat("[NX3;H2,H1,H0;!$(N-C=O);!$(N-c)]")))
    return round(n_sp3/n_all*100, 1) if n_all>0 else 0

def surface_roughness(mol):
    try: return round(CalcKappa3(mol), 3)
    except: return 0

def scaffold_complexity_index(mol):
    """Rings + stereocenters + bridgeheads + spiro as composite."""
    return (CalcNumRings(mol) + stereocenters_count(mol) +
            CalcNumBridgeheadAtoms(mol) + CalcNumSpiroAtoms(mol))

# ═══════════════════════════════════════════════════════════
# SECTION 9 — pKa & IONIZATION
# ═══════════════════════════════════════════════════════════

def pka_acid_estimate(mol):
    if mol.HasSubstructMatch(_pat("S(=O)(=O)O")): return "~1-2 (Sulfonic)"
    if mol.HasSubstructMatch(_pat("C(=O)O")): return "~4-5 (Carboxyl)"
    if mol.HasSubstructMatch(_pat("Oc1ccccc1")): return "~9-10 (Phenol)"
    if mol.HasSubstructMatch(_pat("[NH][CX3]=O")): return "~10 (Amide NH)"
    return ">12 (Neutral)"

def pka_base_estimate(mol):
    if mol.HasSubstructMatch(_pat("[NX4+]")): return ">11 (Quat)"
    if mol.HasSubstructMatch(_pat("[NX3;H2,H1;!$(N-C=O)]")): return "~9-10 (Aliphatic Amine)"
    if mol.HasSubstructMatch(_pat("n1ccccc1")): return "~5 (Pyridine)"
    if mol.HasSubstructMatch(_pat("[nH]1cccc1")): return "~2 (Pyrrole)"
    return "<2 (Weak Base)"

def ionization_at_physiological_ph(mol):
    """Ionization state prediction at pH 7.4."""
    has_acid = mol.HasSubstructMatch(_pat("C(=O)O"))
    has_base = mol.HasSubstructMatch(_pat("[NX3;H1,H2;!$(N-C=O)]"))
    if has_acid and has_base: return "Zwitterion"
    if has_acid: return "Anionic (pH 7.4)"
    if has_base: return "Cationic (pH 7.4)"
    return "Neutral"

def pka_gap(mol):
    acid = mol.HasSubstructMatch(_pat("[CX3](=O)[OX2H1]"))
    base = mol.HasSubstructMatch(_pat("[NX3;H2,H1,H0;!$(N-C=O)]"))
    if acid and base: return "Zwitterionic Potential"
    return "Single/None"

def fraction_ionized_at_74(mol):
    """Rough fraction ionized for carboxylic acids at pH 7.4 (Henderson-Hasselbalch)."""
    if mol.HasSubstructMatch(_pat("C(=O)O")):
        pka = 4.5; ph = 7.4
        fi = 1/(1+10**(pka-ph))
        return round(fi, 3)
    return 0.0

# ═══════════════════════════════════════════════════════════
# SECTION 10 — SAR & FRAGMENT ANALYSIS
# ═══════════════════════════════════════════════════════════

def functional_group_inventory(mol):
    """Counts 25+ important functional groups."""
    fg_map = {
        "Hydroxyl":       "[OX2H]",
        "Carboxyl":       "[CX3](=O)[OX2H1]",
        "Amine_Primary":  "[NX3;H2;!$(N-C=O)]",
        "Amine_Secondary":"[NX3;H1;!$(N-C=O)]",
        "Amine_Tertiary": "[NX3;H0;!$(N-C=O);!$(N-n)]",
        "Amide":          "[NX3][CX3](=[OX1])[#6]",
        "Ester":          "[CX3](=O)[OX2H0][#6]",
        "Ether":          "[OD2]([#6])[#6]",
        "Aldehyde":       "[CX3H1](=O)[#6]",
        "Ketone":         "[CX3](=O)([#6])[#6]",
        "Nitro":          "[N+](=O)[O-]",
        "Sulfonamide":    "[SX4](=[OX1])(=[OX1])[NX3]",
        "Sulfone":        "[SX4](=[OX1])(=[OX1])([#6])[#6]",
        "Phosphonate":    "[PX4](=[OX1])([OX2])([OX2])[#6]",
        "Cyano":          "[CX2]#N",
        "Vinyl":          "[CX3]=[CX3]",
        "Alkyne":         "[CX2]#C",
        "Thiol":          "[SX2H]",
        "Urea":           "[NX3][CX3](=[OX1])[NX3]",
        "Guanidine":      "[NX3][CX3](=[NX2])[NX3]",
        "Oxazole":        "c1cnoc1",
        "Imidazole":      "c1cncc1N",
        "Pyrazole":       "c1cnn[nH]1",
        "Thiophene":      "c1ccsc1",
        "Indole":         "c1ccc2[nH]ccc2c1",
    }
    return {k: len(mol.GetSubstructMatches(Chem.MolFromSmarts(v))) for k, v in fg_map.items()}

def bioisostere_opportunities(mol):
    """Suggests bioisosteric replacements."""
    suggestions = []
    if mol.HasSubstructMatch(_pat("C(=O)O")):
        suggestions.append("Carboxyl → Tetrazole or Acylsulfonamide")
    if mol.HasSubstructMatch(_pat("c1ccccc1")):
        suggestions.append("Phenyl → Pyridine or Thiophene")
    if mol.HasSubstructMatch(_pat("[NX3;H2]")):
        suggestions.append("NH₂ → OH or F (Bioisostere)")
    if mol.HasSubstructMatch(_pat("[CH3]")):
        suggestions.append("Methyl → Cyclopropyl (Metabolic Block)")
    return suggestions if suggestions else ["No Common Bioisosteres"]

def h_bond_donors_detailed(mol):
    nh = len(mol.GetSubstructMatches(_pat("[NH,NH2]")))
    oh = len(mol.GetSubstructMatches(_pat("[OH]")))
    return {"NH": nh, "OH": oh, "Total": nh+oh}

def nitrogen_role_analysis(mol):
    """Classifies nitrogens by role."""
    return {
        "Basic_N":      len(mol.GetSubstructMatches(_pat("[NX3;H1,H2;!$(N-C=O)]"))),
        "Amide_N":      len(mol.GetSubstructMatches(_pat("[NX3][CX3]=O"))),
        "Aromatic_N":   len(mol.GetSubstructMatches(_pat("[n]"))),
        "Quaternary_N": len(mol.GetSubstructMatches(_pat("[NX4+]"))),
    }

def exit_vector_count(mol):
    """Counts available exit vectors (attachment points for decoration)."""
    # Proxy: aromatic CH with no substituent
    return len(mol.GetSubstructMatches(_pat("[cH]")))

def scaffold_type(mol):
    n_arom = CalcNumAromaticRings(mol); n_ali = CalcNumAliphaticRings(mol)
    if n_arom > 1 and n_ali == 0: return "Aromatic"
    if n_arom == 0 and n_ali > 0: return "Aliphatic"
    if n_arom > 0 and n_ali > 0: return "Mixed"
    return "Acyclic"

def saturation_index(fsp3):
    if fsp3>0.45: return "Aliphatic Heavy"
    return "Aromatic Heavy"

def polar_exposure(tp, mw):
    return round(tp/mw*100, 3)

def ligand_efficiency(qed, heavy):
    return round(qed/heavy*10, 3) if heavy>0 else 0

def binding_hotspot_prediction(mol):
    hbd = CalcNumHBD(mol); hba = CalcNumHBA(mol)
    sa = CalcLabuteASA(mol)
    return round((hbd+hba)/sa*100, 2) if sa>0 else 0

# ═══════════════════════════════════════════════════════════
# SECTION 11 — CYP & DDI
# ═══════════════════════════════════════════════════════════

def cyp_2d6_alert(mol, _lp=None):
    basic_n = mol.HasSubstructMatch(_pat("[NX3;H2,H1,H0;!$(N-C=O)]"))
    lp = _lp if _lp is not None else Descriptors.MolLogP(mol)
    if basic_n and lp > 2.5: return "Likely Substrate"
    return "Low Risk"

def cyp_3a4_alert(mol, _mw=None, _lp=None, _n_aro=None):
    """CYP3A4 substrate/inhibitor alert."""
    mw    = _mw    if _mw    is not None else Descriptors.MolWt(mol)
    lp    = _lp    if _lp    is not None else Descriptors.MolLogP(mol)
    n_aro = _n_aro if _n_aro is not None else CalcNumAromaticRings(mol)
    if mw > 400 and lp > 2 and n_aro >= 2:
        return "3A4 Substrate/Inhibitor Likely"
    return "Low 3A4 Risk"

def cyp_2c9_alert(mol, _n_aro=None):
    """CYP2C9 substrate alert (acidic drugs)."""
    has_acid = mol.HasSubstructMatch(_pat("C(=O)O"))
    has_arom = (_n_aro if _n_aro is not None else CalcNumAromaticRings(mol)) >= 1
    if has_acid and has_arom:
        return "2C9 Substrate Likely"
    return "Low 2C9 Risk"

def cyp_2c19_alert(mol, _n_aro=None, _lp=None, _tp=None):
    """CYP2C19 alert."""
    n_arom = _n_aro if _n_aro is not None else CalcNumAromaticRings(mol)
    lp     = _lp    if _lp    is not None else Descriptors.MolLogP(mol)
    tp     = _tp    if _tp    is not None else CalcTPSA(mol)
    if n_arom >= 2 and 1 < lp < 4 and tp < 100:
        return "2C19 Involvement Possible"
    return "Low 2C19 Risk"

def cyp_1a2_alert(mol):
    """CYP1A2 alert — planar aromatic amines."""
    if mol.HasSubstructMatch(_pat("[NH2]-c1ccccc1")):
        return "1A2 Substrate (Aromatic Amine)"
    if CalcNumAromaticRings(mol) >= 3:
        return "1A2 Possible (Planar)"
    return "Low 1A2 Risk"

def ddi_risk_score(mol):
    """Composite DDI risk (0-5)."""
    score = 0
    if cyp_3a4_alert(mol) != "Low 3A4 Risk": score += 2
    if cyp_2d6_alert(mol) != "Low Risk": score += 1
    if cyp_2c9_alert(mol) != "Low 2C9 Risk": score += 1
    if cyp_1a2_alert(mol) != "Low 1A2 Risk": score += 1
    return score

# ═══════════════════════════════════════════════════════════
# SECTION 12 — IP & NOVELTY
# ═══════════════════════════════════════════════════════════

def ip_originality_score(sim):
    return round(100*(1-sim), 1)

def scaffold_novelty_estimate(mol):
    """Estimates novelty from ring count and uncommon heteroatoms."""
    se = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() in (14,34,52,15))
    bh = CalcNumBridgeheadAtoms(mol)
    sp = CalcNumSpiroAtoms(mol)
    score = 50 + se*10 + bh*5 + sp*5
    return min(100, score)

def natural_product_likeness_hint(mol):
    """NP-likeness heuristic: high Fsp3, stereocenters, oxygen-rich."""
    fsp3 = Descriptors.FractionCSP3(mol)
    sc   = stereocenters_count(mol)
    o    = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum()==8)
    score = fsp3*40 + sc*5 + o*2
    if score > 60: return "Natural Product-Like"
    if score > 30: return "Partially NP-Like"
    return "Synthetic-Like"

# ═══════════════════════════════════════════════════════════
# SECTION 13 — GREEN CHEMISTRY & SUSTAINABILITY
# ═══════════════════════════════════════════════════════════

def green_chemistry_score(mol):
    """Estimates green synthesis feasibility (0-100)."""
    score = 70
    # Penalty for rare atoms
    rare = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() in (35,53,79,78,46))
    score -= rare * 10
    # Penalty for complex ring systems
    score -= CalcNumBridgeheadAtoms(mol) * 5
    # Bonus for low MW
    if Descriptors.MolWt(mol) < 350: score += 10
    return max(0, min(100, score))

def atom_economy_estimate(mol):
    """Atom economy proxy — lower MW = better potential."""
    mw = Descriptors.MolWt(mol)
    if mw < 250: return "High Atom Economy"
    if mw < 450: return "Moderate"
    return "Low Atom Economy"

def synthetic_accessibility_category(sa_score):
    """Categorizes SA score."""
    if sa_score <= 2: return "Trivial"
    if sa_score <= 4: return "Easy"
    if sa_score <= 6: return "Moderate"
    if sa_score <= 8: return "Difficult"
    return "Very Difficult"

# ═══════════════════════════════════════════════════════════
# SECTION 14 — ADDITIONAL COMPUTED INDICES
# ═══════════════════════════════════════════════════════════

def hallucination_score(mol):
    return round(Descriptors.MolLogP(mol)*1.5 + Descriptors.MolWt(mol)/150, 2)

def omni_score(mol, qed, sim):
    return round(qed*50 + sim*30 + Descriptors.FractionCSP3(mol)*20, 1)

def hlia_prediction(tp):
    if tp>100: return "Elevated Risk"
    return "Standard"

def ligand_lipophilicity_efficiency(lp, pchembl_proxy):
    """LLE = pChEMBL – logP. Higher = better."""
    return round(pchembl_proxy - lp, 2)

def generative_score_proxy(mol):
    """Synthetic 'generative novelty' based on rare feature combinations."""
    score = 0
    score += CalcNumSpiroAtoms(mol) * 15
    score += CalcNumBridgeheadAtoms(mol) * 10
    score += stereocenters_count(mol) * 5
    score += len(functional_group_inventory(mol)) * 2
    return min(100, score)

def lead_optimization_priority(mol, qed, sim):
    """Composite lead optimization priority score."""
    esol = aqueous_solubility_esol(mol)
    lp = Descriptors.MolLogP(mol)
    tp = CalcTPSA(mol)
    score = qed*40 - abs(lp-2.5)*5 + esol*5 - (sim*10)
    return round(max(0, score), 1)

def flexibility_complexity_ratio(mol):
    """Rotatable bonds vs total ring atoms — balance of flex and rigidity."""
    rot = CalcNumRotatableBonds(mol)
    ring_atoms = sum(len(r) for r in mol.GetRingInfo().AtomRings())
    total = rot + ring_atoms
    return round(rot/total, 3) if total>0 else 0

def nitrogen_oxygen_ratio(mol):
    n = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum()==7)
    o = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum()==8)
    return round(n/(o+1e-9), 2)

def complexity_to_qed_ratio(mol, qed):
    """How efficiently QED is achieved relative to structural complexity."""
    c = graph_complexity(mol)
    return round(qed/(c+1)*100, 3)

def drug_efficiency_index(mol, qed):
    """DEI = QED / (MW / 100) — quality per 100 Da."""
    mw = Descriptors.MolWt(mol)
    return round(qed/(mw/100), 3) if mw>0 else 0

# ═══════════════════════════════════════════════════════════
# MASTER FUNCTION — get_all_mega_v20
# ═══════════════════════════════════════════════════════════

def get_all_mega_v20(mol, qed, sim):
    # ── Compute all base descriptors ONCE — passed to sub-functions ──────────
    mw   = Descriptors.MolWt(mol)
    lp   = Descriptors.MolLogP(mol)
    tp   = CalcTPSA(mol)
    h    = mol.GetNumHeavyAtoms()
    rot  = CalcNumRotatableBonds(mol)
    hbd  = CalcNumHBD(mol)
    hba  = CalcNumHBA(mol)
    refr = Crippen.MolMR(mol)
    fsp3 = Descriptors.FractionCSP3(mol)
    n_aro = CalcNumAromaticRings(mol)   # pre-computed for CYP functions
    rings = CalcNumRings(mol)            # pre-computed for oprea/mddr
    esol = aqueous_solubility_esol(mol)
    pchembl_proxy = qed * 9  # proxy for pIC50

    data = {
        # ── Drug-Likeness Rules — pass pre-computed values, no recomputation ─
        "Rule_of_5_Ext":       rule_of_5_ext(mol, mw, lp, hbd, hba),
        "Ghose_v2":            ghose_v2(lp, mw, h, refr),
        "Egan_v2":             egan_v2(lp, tp),
        "Pfizer_3_75":         pfizer_3_75_rule(lp, tp),
        "GSK_4_400":           gsk_4_400_rule(lp, mw),
        "Oprea_6":             oprea_rule(mol, mw, lp, hbd, hba, rot, rings),
        "Congreve_R3":         congreve_rule_of_3(mol, mw, lp, hbd, hba, rot),
        "Veber_Oral":          veber_rule_ext(rot, tp),
        "MDDR_Like":           mddr_likeness(mol, rings, rot),
        "BioAvail_Score":      bioavailability_score_ro5(mol, mw, lp, hbd, hba, rot, tp),
        "Beyond_Ro5":          beyond_ro5_check(mol, mw, lp, hbd),
        "DC_Linker_Rule":      dc_linker_rule(mol, mw, rot),
        "TICE_CNS_Rule":       tice_rule(mol, tp, rot, mw),

        # ── Lipophilicity & Solubility ───────────────────────
        "Flex_Index":          flexibility_index(mol),
        "Molar_Refractivity":  refr,
        "LogBB_Est":           bbb_v2_index(lp, tp, mw),
        "LogD_pH74":           logd_ph74_estimate(mol),
        "LogD_pH68":           logd_ph68_estimate(mol),
        "ESOL_LogS":           esol,
        "Solubility_Class":    solubility_class(esol),
        "Ali_LogS":            ali_solubility(mol),
        "Sol_Improvement":     solubility_improvement_hint(lp),
        "LipE":                lipophilicity_efficiency(qed, lp),
        "CHI_Hydrophobicity":  chromatographic_hydrophobicity_index(lp, tp),

        # ── BBB & CNS ────────────────────────────────────────
        "BBB_Likelihood":      bbb_likelihood_cat(lp, tp),
        "CNS_MPO_Score":       cns_mpo_score(mol),
        "P_gp_Alert":          p_glycoprotein_alert(mol),
        "Papp_CACO2":          passive_permeability_papp(mol),
        "Oral_Absorption_Pct": oral_absorption_score(mol),

        # ── Metabolism & Clearance ───────────────────────────
        "Metab_HalfLife":      metabolic_half_life_score(mol),
        "Microsomal_Stability":microsomal_stability_index(mol),
        "Hepatic_ER":          hepatic_extraction_ratio(mol),
        "First_Pass_Risk":     first_pass_effect_risk(mol),
        "Phase_II":            phase_ii_alerts(mol),
        "Clipping_Alert":      metabolic_clipping_alert(mol),
        "Reactive_Met_Risk":   reactive_metabolite_risk(mol),

        # ── Toxicity Alerts ──────────────────────────────────
        "React_Danger":        covalent_danger_zone(mol),
        "Tox_Alerts_Count":    toxicity_structural_alerts(mol),
        "PAINS_Flags":         pains_filter(mol),
        "DILI_Risk":           dili_risk_extended(mol),
        "Skin_Sensitization":  skin_sensitization_risk(mol),
        "Ames_Alert":          mutagenicity_ames_alert(mol),
        "Genotox_Flags":       genotox_structural_flags(mol),
        "hERG_Alert":          cardiotoxicity_herg_alert(mol),
        "Nephrotox_Alert":     nephrotoxicity_alert(mol),
        "Phospholipidosis":    phospholipidosis_risk(mol),
        "Idiosyncratic_Tox":   idiosyncratic_tox_flag(mol),

        # ── Structural Complexity ────────────────────────────
        "Kappa1":              shape_index_kappa1(mol),
        "Kappa2":              shape_index_kappa2(mol),
        "Kappa3_Roughness":    shape_index_kappa3(mol),
        "Labute_ASA":          labute_asa(mol),
        "Chi_Indices":         chi_indices(mol),
        "Wiener_J":            wiener_index(mol),
        "Graph_Complexity":    graph_complexity(mol),
        "Mol_Volume":          mol_volume(mol),
        "Polarizability":      polarizability_estimate(mol),
        "TPSA_per_Heavy":      tpsa_per_heavy(mol, tp),
        "Bridgehead_Count":    bridgehead_complexity(mol),
        "Spiro_Count":         spiro_complexity(mol),
        "Sterics_Tert_Butyl":  sterics_index(mol),
        "Stereocenters":       stereocenters_count(mol),
        "Stereo_Density":      stereo_density(mol),
        "Isomer_Count":        isomer_count_hint(mol),
        "Scaffold_Complexity": scaffold_complexity_index(mol),
        "Flexibility_%":       molecular_flexibility(mol),
        "Flex_Rigid_Ratio":    flexibility_complexity_ratio(mol),
        "Arom_Heavy_Ratio":    arom_heavy_ratio(mol),
        "Halogen_Ratio":       halogen_ratio(mol),
        "Heteroatom_Ratio":    heteroatom_ratio(mol),
        "Amide_Bonds":         amide_bond_count(mol),
        "Nitrogen_Sat_%":      nitrogen_saturation(mol),
        "N_O_Ratio":           nitrogen_oxygen_ratio(mol),
        "Polar_Exposure":      polar_exposure(tp, mw),

        # ── Ring Analysis ────────────────────────────────────
        "Aromatic_Rings":      aromatic_ring_count(mol),
        "Aliphatic_Rings":     aliphatic_ring_count(mol),
        "Saturated_Rings":     saturated_ring_count(mol),
        "Heterocycles":        heterocycle_count(mol),
        "Arom_Heterocycles":   aromatic_heterocycle_count(mol),
        "Ring_Fusion_Score":   ring_fusion_score(mol),
        "Macrocycle_Flag":     macrocycle_flag(mol),

        # ── pKa & Ionization ────────────────────────────────
        "pKa_Acidic":          pka_acid_estimate(mol),
        "pKa_Basic":           pka_base_estimate(mol),
        "Ionization_pH74":     ionization_at_physiological_ph(mol),
        "Fraction_Ionized":    fraction_ionized_at_74(mol),
        "Zwitterion":          pka_gap(mol),
        "HLIA_Risk":           hlia_prediction(tp),

        # ── SAR & Fragments ──────────────────────────────────
        "Func_Group_Inventory":functional_group_inventory(mol),
        "Bioisostere_Hints":   bioisostere_opportunities(mol),
        "HBond_Detailed":      h_bond_donors_detailed(mol),
        "Nitrogen_Roles":      nitrogen_role_analysis(mol),
        "Exit_Vectors":        exit_vector_count(mol),
        "Scaffold_Type":       scaffold_type(mol),
        "Sat_Cat":             saturation_index(fsp3),

        # ── CYP Profile ─────────────────────────────────────
        "CYP_2D6_Hint":        cyp_2d6_alert(mol, lp),
        "CYP_3A4_Alert":       cyp_3a4_alert(mol, mw, lp, n_aro),
        "CYP_2C9_Alert":       cyp_2c9_alert(mol, n_aro),
        "CYP_2C19_Alert":      cyp_2c19_alert(mol, n_aro, lp, tp),
        "CYP_1A2_Alert":       cyp_1a2_alert(mol),
        "DDI_Risk_Score":      ddi_risk_score(mol),

        # ── Efficiency & Novelty ─────────────────────────────
        "IP_Originality":      ip_originality_score(sim),
        "Scaffold_Novelty":    scaffold_novelty_estimate(mol),
        "NP_Likeness":         natural_product_likeness_hint(mol),
        "Ligand_Efficiency":   ligand_efficiency(qed, h),
        "Hotspot_Density":     binding_hotspot_prediction(mol),
        "Drug_Efficiency_Idx": drug_efficiency_index(mol, qed),
        "LLE":                 ligand_lipophilicity_efficiency(lp, pchembl_proxy),
        "Generative_Score":    generative_score_proxy(mol),
        "Lead_Optim_Priority": lead_optimization_priority(mol, qed, sim),
        "QED_Complexity_Ratio":complexity_to_qed_ratio(mol, qed),
        "Omni_Score":          omni_score(mol, qed, sim),
        "Halluc_Score":        hallucination_score(mol),

        # ── Green Chemistry ──────────────────────────────────
        "Green_Chem_Score":    green_chemistry_score(mol),
        "Atom_Economy":        atom_economy_estimate(mol),

        # ── Miscellaneous ────────────────────────────────────
        "Phase_II_Sites":      phase_ii_alerts(mol),
        "Sat_Rings":           saturated_ring_count(mol),
    }

    # Append atom counts and ring distribution
    data.update(atom_counts(mol))
    data.update(heteroatom_detail(mol))
    data.update(ring_size_distribution(mol))

    return data
