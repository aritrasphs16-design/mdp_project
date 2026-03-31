"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   CHEMOFILTER — DRUG DISCOVERY EXTENDED COLUMNS (SwissADME-Grade)           ║
║   HBD/HBA · RotBonds · ADMET · Structural · Lead Optimization · Badges     ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from rdkit import Chem
from rdkit.Chem import Descriptors, QED, Crippen, AllChem
from rdkit.Chem.rdMolDescriptors import (
    CalcNumHBD, CalcNumHBA, CalcNumRotatableBonds,
    CalcNumRings, CalcNumAromaticRings, CalcNumAliphaticRings,
    CalcTPSA, CalcLabuteASA, CalcFractionCSP3,
    CalcNumBridgeheadAtoms, CalcNumSpiroAtoms,
    CalcNumHeteroatoms, CalcNumAmideBonds,
    CalcNumHeterocycles
)
import math
import random


# ═══════════════════════════════════════════════════════════════════
# 1. DRUG-DISCOVERY COLUMNS (Lipinski extended)
# ═══════════════════════════════════════════════════════════════════

def get_extended_descriptors(mol):
    """Compute all extended drug-discovery columns for a molecule."""
    mw = Descriptors.MolWt(mol)
    logp = Descriptors.MolLogP(mol)
    tpsa = CalcTPSA(mol)
    hbd = CalcNumHBD(mol)
    hba = CalcNumHBA(mol)
    rot_bonds = CalcNumRotatableBonds(mol)
    ring_count = CalcNumRings(mol)
    arom_ring_count = CalcNumAromaticRings(mol)
    heavy_atom_count = mol.GetNumHeavyAtoms()
    fsp3 = CalcFractionCSP3(mol)
    frac_aromatic = sum(1 for a in mol.GetAtoms() if a.GetIsAromatic()) / heavy_atom_count if heavy_atom_count > 0 else 0

    # Lipinski violations
    lip_violations = sum([mw > 500, logp > 5, hbd > 5, hba > 10])

    # Veber Rule Score: Pass if RotBonds ≤ 10 AND TPSA ≤ 140
    veber_score = 1 if (rot_bonds <= 10 and tpsa <= 140) else 0

    # Ghose Filter Score: MW 160-480, LogP -0.4-5.6, Atoms 20-70, MR 40-130
    mr = Crippen.MolMR(mol)
    ghose_pass = sum([
        160 <= mw <= 480,
        -0.4 <= logp <= 5.6,
        20 <= heavy_atom_count <= 70,
        40 <= mr <= 130
    ])

    return {
        "HBD": hbd,
        "HBA": hba,
        "Rotatable_Bonds": rot_bonds,
        "Ring_Count": ring_count,
        "Aromatic_Ring_Count": arom_ring_count,
        "Heavy_Atom_Count": heavy_atom_count,
        "Fraction_Aromatic": round(frac_aromatic, 3),
        "Lipinski_Violations": lip_violations,
        "Veber_Rule_Score": veber_score,
        "Ghose_Filter_Score": f"{ghose_pass}/4",
    }


# ═══════════════════════════════════════════════════════════════════
# 2. ADMET PREDICTION COLUMNS
# ═══════════════════════════════════════════════════════════════════

def get_admet_predictions(mol):
    """SwissADME / Schrödinger style ADMET predictions."""
    logp = Descriptors.MolLogP(mol)
    tpsa = CalcTPSA(mol)
    mw = Descriptors.MolWt(mol)
    hbd = CalcNumHBD(mol)
    hba = CalcNumHBA(mol)
    rot = CalcNumRotatableBonds(mol)
    heavy = mol.GetNumHeavyAtoms()

    # Solubility (LogS) — ESOL estimate
    arom_rings = CalcNumAromaticRings(mol)
    log_s = round(0.16 - 0.63 * logp - 0.0062 * mw + 0.066 * rot - 0.74 * arom_rings, 2)
    if log_s > 0:
        sol_class = "Highly Soluble"
    elif log_s > -2:
        sol_class = "Soluble"
    elif log_s > -4:
        sol_class = "Moderate"
    elif log_s > -6:
        sol_class = "Poorly Soluble"
    else:
        sol_class = "Insoluble"

    # Human Intestinal Absorption
    hia = "High" if tpsa < 140 else ("Moderate" if tpsa < 200 else "Low")

    # Blood Brain Barrier
    bbb = "Yes" if (tpsa < 79 and logp > 0 and logp < 5 and mw < 450) else "No"

    # CYP450 Interaction Risk
    cyp_risk = 0
    if mol.HasSubstructMatch(Chem.MolFromSmarts("[NX3;H1,H2;!$(N-C=O)]")): cyp_risk += 1
    if CalcNumAromaticRings(mol) >= 2: cyp_risk += 1
    if logp > 3: cyp_risk += 1
    if mw > 400: cyp_risk += 1
    cyp_level = "High" if cyp_risk >= 3 else ("Medium" if cyp_risk >= 2 else "Low")

    # Plasma Protein Binding
    if logp > 4.5:
        ppb = ">95%"
    elif logp > 3.0:
        ppb = "85-95%"
    elif logp > 1.0:
        ppb = "50-85%"
    else:
        ppb = "<50%"

    # Clearance Prediction
    if logp > 4 or mw > 500:
        clearance = "High (Hepatic)"
    elif logp < 1 and mw < 300:
        clearance = "High (Renal)"
    else:
        clearance = "Moderate"

    # Half Life Prediction
    if logp > 4 and mw > 400:
        half_life = "Long (>12h)"
    elif logp < 1:
        half_life = "Short (<4h)"
    else:
        half_life = "Medium (4-12h)"

    # Bioavailability Score (0-1)
    bio_score = 0.0
    if mw <= 500: bio_score += 0.2
    if logp <= 5: bio_score += 0.2
    if hbd <= 5: bio_score += 0.2
    if hba <= 10: bio_score += 0.2
    if rot <= 10: bio_score += 0.1
    if tpsa <= 140: bio_score += 0.1

    # Toxicity Risk Score
    tox_alerts = 0
    tox_patterns = [
        "[N+](=O)[O-]", "c1ccccc1N", "C1OC1", "C=CC=O",
        "[NX2]=O", "[As,Sb,Hg,Pb,Cd]", "[NX3][NX3]"
    ]
    for pat in tox_patterns:
        p = Chem.MolFromSmarts(pat)
        if p and mol.HasSubstructMatch(p):
            tox_alerts += 1
    tox_level = "High" if tox_alerts >= 3 else ("Medium" if tox_alerts >= 1 else "Low")

    # Mutagenicity Risk
    mutagen_alerts = []
    mutagen_patterns = {
        "Nitro": "[N+](=O)[O-]",
        "Aromatic Amine": "c1ccccc1[NH2]",
        "Nitroso": "[NX2]=O",
        "Epoxide": "C1OC1",
        "Hydrazine": "[NX3][NX3]",
    }
    for name, sma in mutagen_patterns.items():
        p = Chem.MolFromSmarts(sma)
        if p and mol.HasSubstructMatch(p):
            mutagen_alerts.append(name)
    mutagen_risk = "High" if len(mutagen_alerts) >= 2 else ("Medium" if len(mutagen_alerts) >= 1 else "Low")

    return {
        "LogS_ESOL": log_s,
        "Solubility_Class": sol_class,
        "HIA": hia,
        "BBB_Penetration": bbb,
        "CYP450_Risk": cyp_level,
        "Plasma_Protein_Binding": ppb,
        "Clearance": clearance,
        "Half_Life": half_life,
        "Bioavailability_Score": round(bio_score, 2),
        "Toxicity_Risk": tox_level,
        "Mutagenicity_Risk": mutagen_risk,
        "Mutagenicity_Alerts": mutagen_alerts,
    }


# ═══════════════════════════════════════════════════════════════════
# 3. STRUCTURAL CHEMISTRY METRICS
# ═══════════════════════════════════════════════════════════════════

def get_structural_metrics(mol):
    """Advanced structural complexity indicators."""
    # Chiral Centers Count
    chiral = len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))

    # Ring Strain Score (heuristic: small rings + bridgeheads)
    ri = mol.GetRingInfo()
    small_rings = sum(1 for r in ri.AtomRings() if len(r) <= 4)
    bridgeheads = CalcNumBridgeheadAtoms(mol)
    ring_strain = small_rings * 3 + bridgeheads * 2

    # Topological Complexity (Bertz-inspired)
    try:
        from rdkit.Chem import GraphDescriptors
        topo_complexity = round(GraphDescriptors.BertzCT(mol), 1)
    except:
        topo_complexity = 0

    # Fragment Diversity Score
    fg_patterns = [
        "[OX2H]", "[CX3](=O)[OX2H1]", "[NX3;H2;!$(N-C=O)]",
        "[NX3][CX3](=[OX1])", "[CX3](=O)[OX2H0]", "[F,Cl,Br,I]",
        "[SX4](=[OX1])(=[OX1])", "[CX2]#N", "c1ccccc1", "[nH]"
    ]
    frag_count = sum(1 for p in fg_patterns
                     if Chem.MolFromSmarts(p) and mol.HasSubstructMatch(Chem.MolFromSmarts(p)))
    frag_diversity = round(frag_count / len(fg_patterns), 2)

    # Molecular Symmetry Score (heuristic)
    canon = Chem.MolToSmiles(mol)
    atoms = mol.GetNumAtoms()
    unique_types = len(set(a.GetAtomicNum() for a in mol.GetAtoms()))
    symmetry = round((1 - unique_types / max(atoms, 1)) * 100, 1)

    # 3D Shape Index (Kappa3)
    try:
        from rdkit.Chem.rdMolDescriptors import CalcKappa3
        shape_3d = round(CalcKappa3(mol), 3)
    except:
        shape_3d = 0

    return {
        "Chiral_Centers": chiral,
        "Ring_Strain_Score": ring_strain,
        "Topological_Complexity": topo_complexity,
        "Fragment_Diversity": frag_diversity,
        "Molecular_Symmetry": symmetry,
        "Shape_Index_3D": shape_3d,
    }


# ═══════════════════════════════════════════════════════════════════
# 4. ADVANCED LEAD OPTIMIZATION INDICATORS
# ═══════════════════════════════════════════════════════════════════

def get_lead_optimization(mol, qed_val=None):
    """Indicators to help chemists decide which compound to optimize."""
    mw = Descriptors.MolWt(mol)
    logp = Descriptors.MolLogP(mol)
    tpsa = CalcTPSA(mol)
    hbd = CalcNumHBD(mol)
    hba = CalcNumHBA(mol)
    rot = CalcNumRotatableBonds(mol)
    heavy = mol.GetNumHeavyAtoms()
    fsp3 = CalcFractionCSP3(mol)
    qed = qed_val if qed_val is not None else QED.qed(mol)

    # Lead Optimization Potential (0-100)
    lop = 100
    if mw > 500: lop -= 15
    if logp > 5: lop -= 15
    if logp < -1: lop -= 10
    if hbd > 5: lop -= 10
    if hba > 10: lop -= 10
    if rot > 10: lop -= 10
    if tpsa > 140: lop -= 10
    if fsp3 < 0.2: lop -= 5
    lop = max(0, min(100, lop))

    # Fragment Efficiency (QED / heavy atoms * 100)
    frag_eff = round(qed / heavy * 100, 2) if heavy > 0 else 0

    # Ligand Efficiency (proxy: QED normalized by mol weight)
    lig_eff = round(qed / (mw / 100), 3) if mw > 0 else 0

    # Lipophilic Efficiency (LipE = pActivity - LogP, proxy with QED)
    lip_eff = round(qed * 10 - logp, 2)

    # Synthetic Route Difficulty (based on SA Score concepts)
    rings = CalcNumRings(mol)
    chiral = len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))
    synth_difficulty = "Easy" if (rings <= 2 and chiral <= 1 and mw < 350) \
        else ("Hard" if (rings >= 5 or chiral >= 3 or mw > 700) else "Moderate")

    # Optimization Priority Score (composite 0-100)
    opt_priority = round(lop * 0.3 + qed * 30 + frag_eff * 0.2 + (10 - min(logp, 10)) * 2, 1)
    opt_priority = max(0, min(100, opt_priority))

    return {
        "Lead_Optimization_Potential": lop,
        "Fragment_Efficiency": frag_eff,
        "Ligand_Efficiency": lig_eff,
        "Lipophilic_Efficiency": lip_eff,
        "Synthetic_Difficulty": synth_difficulty,
        "Optimization_Priority": opt_priority,
    }


# ═══════════════════════════════════════════════════════════════════
# 5. VISUAL INDICATORS & BADGES
# ═══════════════════════════════════════════════════════════════════

def get_drug_likeness_badge(mol):
    """Classify compound into drug-like categories."""
    mw = Descriptors.MolWt(mol)
    logp = Descriptors.MolLogP(mol)
    hbd = CalcNumHBD(mol)
    hba = CalcNumHBA(mol)
    rot = CalcNumRotatableBonds(mol)
    tpsa = CalcTPSA(mol)
    lip = sum([mw > 500, logp > 5, hbd > 5, hba > 10])

    # Fragment-like: MW<300, LogP<3, HBD≤3, HBA≤3, RotBonds≤3
    if mw < 300 and logp < 3 and hbd <= 3 and hba <= 3 and rot <= 3:
        return "Fragment-like", "#a78bfa"
    # Lead-like: MW<350, LogP<3, HBD<3
    if mw < 350 and logp < 3 and hbd < 3:
        return "Lead-like", "#38bdf8"
    # Drug-like: Lipinski compliant
    if lip <= 1:
        return "Drug-like", "#34d399"
    # Poor candidate
    return "Poor candidate", "#f87171"


def get_property_warnings(mol):
    """Generate warning icons for problematic properties."""
    warnings = []
    logp = Descriptors.MolLogP(mol)
    tpsa = CalcTPSA(mol)
    mw = Descriptors.MolWt(mol)
    hbd = CalcNumHBD(mol)

    if logp > 5:
        warnings.append("⚠ high logP")
    if logp < -2:
        warnings.append("⚠ very low logP")
    if tpsa > 140:
        warnings.append("⚠ high TPSA")
    if mw > 500:
        warnings.append("⚠ high MW")
    if hbd > 5:
        warnings.append("⚠ high HBD")

    # Solubility check
    rot = CalcNumRotatableBonds(mol)
    arom = CalcNumAromaticRings(mol)
    log_s = 0.16 - 0.63 * logp - 0.0062 * mw + 0.066 * rot - 0.74 * arom
    if log_s < -4:
        warnings.append("⚠ poor solubility")

    # hERG
    basic_n = mol.HasSubstructMatch(Chem.MolFromSmarts("[NX3;H1,H2;!$(N-C=O)]"))
    if basic_n and logp > 3 and CalcNumAromaticRings(mol) >= 2:
        warnings.append("⚠ hERG risk")

    return warnings


def heatmap_color(value, low, mid, high):
    """Return green/yellow/red CSS color based on value thresholds."""
    if value <= low:
        return "#34d399"  # green
    elif value <= mid:
        return "#fbbf24"  # yellow
    else:
        return "#f87171"  # red


def get_radar_data(mol):
    """Get normalized 0-1 data for radar chart (Lipophilicity, Size, Polarity, Flexibility, Solubility)."""
    logp = Descriptors.MolLogP(mol)
    mw = Descriptors.MolWt(mol)
    tpsa = CalcTPSA(mol)
    rot = CalcNumRotatableBonds(mol)
    arom = CalcNumAromaticRings(mol)
    log_s = 0.16 - 0.63 * logp - 0.0062 * mw + 0.066 * rot - 0.74 * arom

    lipophilicity = min(1.0, max(0, logp / 7))
    size = min(1.0, max(0, mw / 800))
    polarity = min(1.0, max(0, tpsa / 200))
    flexibility = min(1.0, max(0, rot / 15))
    solubility = min(1.0, max(0, (log_s + 6) / 8))

    return {
        "Lipophilicity": round(lipophilicity, 3),
        "Size": round(size, 3),
        "Polarity": round(polarity, 3),
        "Flexibility": round(flexibility, 3),
        "Solubility": round(solubility, 3),
    }


# ═══════════════════════════════════════════════════════════════════
# 6. COMPREHENSIVE ANALYSIS (ONE-CALL)
# ═══════════════════════════════════════════════════════════════════

def get_full_extended_analysis(mol, qed_val=None):
    """Run ALL extended analyses in one call. Returns merged dict."""
    result = {}
    result.update(get_extended_descriptors(mol))
    result.update(get_admet_predictions(mol))
    result.update(get_structural_metrics(mol))
    result.update(get_lead_optimization(mol, qed_val))

    badge, badge_color = get_drug_likeness_badge(mol)
    result["Drug_Likeness_Badge"] = badge
    result["Badge_Color"] = badge_color
    result["Property_Warnings"] = get_property_warnings(mol)
    result["Radar_Data"] = get_radar_data(mol)

    return result
