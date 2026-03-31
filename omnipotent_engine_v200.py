
from rdkit import Chem
from rdkit.Chem import Descriptors, rdMolDescriptors, AllChem, DataStructs
from rdkit.Chem.rdFingerprintGenerator import GetMorganGenerator as _MorganGen
_MORGAN_GEN = _MorganGen(radius=2, fpSize=2048)  # module-level singleton
import omnipotent_reactivity_db as odb
import master_drug_atlas as mda
import math

class SingularityEngineV200:
    """
    The Ultimate Intelligence Layer (v200). 
    Integrates Metabolic Simulation, Eco-Tox, and Master Atlas Anchoring.
    """
    
    def __init__(self):
        self.metabolic_db = odb.get_metabolic_db()
        self.ddi_db = odb.get_ddi_alerts()
        self.eco_db = odb.get_eco_tox()
        self.rare_db = odb.get_rare_groups()
        self.atlas = mda.get_master_atlas()
        
    def calculate_ligand_efficiency(self, mol, potency_ic50_nm=100):
        """Calculates LE, LLE, and FQ."""
        # Potency to DeltaG: DeltaG = RT ln(Kd)
        # Simplified: pIC50 = -log10(IC50)
        pic50 = -math.log10(potency_ic50_nm * 1e-9)
        ha = mol.GetNumHeavyAtoms()
        lp = Descriptors.MolLogP(mol)
        
        le = (1.37 * pic50) / ha if ha > 0 else 0
        lle = pic50 - lp
        fq = le / (0.01 * ha + 0.3) if ha > 0 else 0 # Fit Quality
        
        return {
            "LE": round(le, 2),
            "LLE": round(lle, 2),
            "Fit_Quality": round(fq, 2)
        }

    def simulate_metabolism(self, mol):
        """Predicts potential metabolic transformations and hazards."""
        sites = []
        for p in self.metabolic_db:
            pat = Chem.MolFromSmarts(p["pattern"])
            if pat and mol.HasSubstructMatch(pat):
                sites.append({
                    "Transformation": p["name"],
                    "Result": p["product_hint"]
                })
        return sites[:8] # Return top 8 matches

    def cross_atlas_similarity_stats(self, mol):
        """Measures distance to the closest drug in the Master Atlas."""
        target_fp = _MORGAN_GEN.GetFingerprint(mol)
        max_sim = 0
        
        for drug in self.atlas:
            d_mol = Chem.MolFromSmiles(drug[1])
            if d_mol:
                d_fp = _MORGAN_GEN.GetFingerprint(d_mol)
                sim = DataStructs.TanimotoSimilarity(target_fp, d_fp)
                if sim > max_sim: max_sim = sim
        
        return round(max_sim * 100, 1)

    def eco_toxicity_deep_scan(self, mol):
        """Scans for environmental and persistence hazards."""
        hits = []
        for e in self.eco_db:
             pat = Chem.MolFromSmarts(e["smarts"])
             if pat and mol.HasSubstructMatch(pat):
                 hits.append({"Danger": e["name"], "Level": e["level"]})
        return hits

    def rare_functional_scout(self, mol):
        """Identifies chemically unique or difficult groups."""
        found = []
        for name, smarts in self.rare_db.items():
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                found.append(name)
        return found

    def binding_site_volume_est(self, mol):
        """Heuristic volume estimation (A^3)."""
        mw = Descriptors.MolWt(mol)
        ha = mol.GetNumHeavyAtoms()
        # V = 1.2 * MW / density (approx 1.2)
        return round(mw * 1.5, 1)

    def analyze_v200(self, mol, base_res):
        le_stats = self.calculate_ligand_efficiency(mol)
        metabolites = self.simulate_metabolism(mol)
        atlas_sim = self.cross_atlas_similarity_stats(mol)
        eco_scan = self.eco_toxicity_deep_scan(mol)
        rares = self.rare_functional_scout(mol)
        vol = self.binding_site_volume_est(mol)
        
        # Singularity Score (0-100)
        # Based on LE, Atlas Sim, and Toxicity absence
        tox_penalty = len(eco_scan) * 20
        sing_score = min(100, max(0, (le_stats["LE"] * 100) + (atlas_sim / 2) - tox_penalty))
        
        return {
            "Singularity_Score": round(sing_score, 1),
            "LE_Metrics": le_stats,
            "Metabolic_Sim": metabolites,
            "Atlas_Distance": atlas_sim,
            "Eco_Tox": eco_scan,
            "Rare_Groups": rares,
            "Heuristic_Volume": vol,
            "Status": "READY" if sing_score > 40 else "UNSTABLE"
        }

_V200_INSTANCE = None

def get_v200_engine():
    """Return a module-level singleton — data loaded only once."""
    global _V200_INSTANCE
    if _V200_INSTANCE is None:
        _V200_INSTANCE = SingularityEngineV200()
    return _V200_INSTANCE
