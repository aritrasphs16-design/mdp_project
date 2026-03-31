
from rdkit import Chem
from rdkit.Chem import Descriptors, rdMolDescriptors
import math

# ── Pre-compiled SMARTS patterns — compiled once at module import ─────────────
_PAT_BASIC_N       = Chem.MolFromSmarts("[NX3;H2,H1,H0;!$(N-C=O)]")
_PAT_ANILINE       = Chem.MolFromSmarts("[NX3;H2][cX3]")
_PAT_THIOAMIDE     = Chem.MolFromSmarts("[NX3][CX3](=S)")
_PAT_PAH           = Chem.MolFromSmarts("c1ccc2c3c(ccc2c1)cccc3")
_PAT_NITRO         = Chem.MolFromSmarts("[N+](=O)[O-]")
_PAT_MICHAEL_ACC   = Chem.MolFromSmarts("[CX3]=[CX3][CX3]=O")
_PAT_ACID          = Chem.MolFromSmarts("[CX3](=O)[OX2H1]")
_PAT_HYDRAZINE     = Chem.MolFromSmarts("[NX3][NX3]")
# ─────────────────────────────────────────────────────────────────────────────

def chembl_likeness(mol):
    """How typical is this molecule compared to bioactive ChEMBL compounds."""
    mw = Descriptors.MolWt(mol)
    lp = Descriptors.MolLogP(mol)
    rings = rdMolDescriptors.CalcNumRings(mol)
    score = 0
    if 250 <= mw <= 450: score += 40
    if 1.0 <= lp <= 4.0: score += 40
    if 2 <= rings <= 4: score += 20
    return score

def fsp3_target(fsp3):
    """Golden Range for Fsp3 is 0.35 - 0.55."""
    if 0.35 <= fsp3 <= 0.55: return "Optimal ✨"
    if fsp3 > 0.55: return "High (Complex)"
    return "Low (Flat)"

def muegge_filter(mol):
    """Muegge (Bayer) pharmacophore point filter."""
    mw = Descriptors.MolWt(mol)
    lp = Descriptors.MolLogP(mol)
    tp = Descriptors.TPSA(mol)
    rings = rdMolDescriptors.CalcNumRings(mol)
    heavy = mol.GetNumHeavyAtoms()
    hbd = rdMolDescriptors.CalcNumHBD(mol)
    hba = rdMolDescriptors.CalcNumHBA(mol)
    
    violations = 0
    if not (200 <= mw <= 600): violations += 1
    if not (-2 <= lp <= 5): violations += 1
    if tp > 150: violations += 1
    if rings > 7: violations += 1
    if heavy <= 4: violations += 1
    if hbd > 5: violations += 1
    if hba > 10: violations += 1
    return "Pass" if violations == 0 else f"Fail ({violations} violations)"

def caco2_perm(lp, tp):
    """Heuristic Caco-2 permeability estimate."""
    if lp > 1.5 and tp < 70: return "High"
    if lp > 0 and tp < 120: return "Moderate"
    return "Low"

def pgp_substrate_alert(mol):
    """Alert for P-glycoprotein efflux."""
    mw = Descriptors.MolWt(mol)
    has_basic_n = mol.HasSubstructMatch(_PAT_BASIC_N)
    if mw > 450 and has_basic_n: return "High Risk"
    return "Low Risk"

def skin_perm_logkp(lp, mw):
    """Potts and Guy equation for skin permeability."""
    return round(0.71 * lp - 0.0061 * mw - 6.3, 2)

def dili_risk(mol):
    """Drug-Induced Liver Injury risk alerts."""
    hits = []
    if mol.HasSubstructMatch(_PAT_ANILINE):   hits.append("Aniline")
    if mol.HasSubstructMatch(_PAT_THIOAMIDE): hits.append("Thioamide")
    if mol.HasSubstructMatch(_PAT_PAH):       hits.append("PAH")
    if mol.HasSubstructMatch(_PAT_NITRO):     hits.append("Nitro Group")
    return "Concerns: " + ", ".join(hits) if hits else "Safe Profile"

def phospholipidosis_risk(mol, lp):
    """Cationic Amphiphilic Drug (CAD) alert."""
    basic_n = len(mol.GetSubstructMatches(_PAT_BASIC_N))
    if basic_n >= 1 and lp > 3: return "Potential Risk"
    return "No Risk"

def vd_prediction(lp):
    """Volume of Distribution estimate (L/kg)."""
    # High LogP increases Vd
    if lp > 4: return "High (> 5 L/kg)"
    if lp > 2: return "Medium (1-5 L/kg)"
    return "Low (< 1 L/kg)"

def half_life_cat(mw, lp):
    """Half-life category heuristic."""
    if lp > 3 and mw > 400: return "Long (> 12h)"
    if lp < 1: return "Short (< 4h)"
    return "Moderate (4-12h)"

def oral_absorption(hia, bbb):
    """Human Oral Absorption % estimate."""
    if hia and bbb: return "High (> 80%)"
    if hia: return "Moderate (50-80%)"
    return "Low (< 30%)"

def scaffold_type(mol):
    """Bemis-Murcko scaffold classification."""
    try:
        from rdkit.Chem.Scaffolds import MurckoScaffold
        scaff = MurckoScaffold.GetScaffoldForMol(mol)
        return Chem.MolToSmiles(scaff) if scaff else "Linear/Small"
    except: return "Unknown"

def ring_complexity(mol):
    """Analysis of ring system architecture."""
    info = mol.GetRingInfo()
    num_rings = info.NumRings()
    if num_rings == 0: return "Acyclic"
    
    # Check for spiro
    spiro = rdMolDescriptors.CalcNumSpiroAtoms(mol)
    # Check for bridgehead
    bridge = rdMolDescriptors.CalcNumBridgeheadAtoms(mol)
    
    if spiro > 0 and bridge > 0: return "Hyper-Complex (Spiro-Bridged)"
    if spiro > 0: return "Spirocentric"
    if bridge > 0: return "Bridged/Fused"
    return "Simple Cyclic"

def stereo_density(mol):
    """Stereocenters per heavy atom."""
    centers = len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))
    heavy = mol.GetNumHeavyAtoms()
    return round((centers / heavy * 100), 1) if heavy > 0 else 0

def hbond_balance(hbd, hba):
    """Ratio of H-bond donors to acceptors."""
    if hba == 0: return hbd
    return round(hbd / hba, 2)

def nephrotox_index(lp, sa):
    """Heuristic for potential kidney stress."""
    if lp < 0 and sa > 4: return "Elevated"
    return "Normal"

def skin_sensitization(mol):
    """Checks for electrophilic skin sensitisers."""
    if mol.HasSubstructMatch(_PAT_MICHAEL_ACC): return "Potential Sensitiser"
    return "Clear"

def logp_logd_gap(mol, lp):
    """Highlights ionizable species."""
    acid = mol.HasSubstructMatch(_PAT_ACID)
    base = mol.HasSubstructMatch(_PAT_BASIC_N)
    if acid or base: return "High (Ionizable)"
    return "Low (Neutral)"

def bpp_ratio(lp):
    """Blood-to-Plasma Partition estimate."""
    if lp > 3: return "High Entry (> 1.2)"
    return "Standard (0.7 - 1.0)"

def genotox_breslow(mol):
    """Alerts based on Breslow's rules for mutagenicity."""
    if mol.HasSubstructMatch(_PAT_NITRO) or mol.HasSubstructMatch(_PAT_HYDRAZINE): return "Geno-Tox Alert"
    return "Low Risk"
