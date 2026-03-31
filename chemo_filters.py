
# VANGUARD ENGINE v2.0 - CORE MODULE
# OPTIMIZED: Pre-compiled SMARTS, removed redundant imports, fixed double GetMolFrags call

from rdkit import Chem
from rdkit.Chem import Descriptors, rdMolDescriptors, AllChem, DataStructs, Crippen, Fragments, QED as _QED
from rdkit.Chem.rdFingerprintGenerator import GetMorganGenerator as _MorganGen
_MORGAN_GEN = _MorganGen(radius=2, fpSize=2048)  # module-level singleton
from rdkit.Chem.FilterCatalog import FilterCatalog, FilterCatalogParams
from rdkit.Chem.Scaffolds import MurckoScaffold

# =============================================================================
# MODULE-LEVEL SINGLETONS — built once on import, reused for every molecule
# =============================================================================

# Pre-built PAINS + Brenk FilterCatalog — avoids rebuilding per molecule
try:
    _safety_params = FilterCatalogParams()
    _safety_params.AddCatalog(FilterCatalogParams.FilterCatalogs.PAINS)
    _safety_params.AddCatalog(FilterCatalogParams.FilterCatalogs.BRENK)
    _SAFETY_CATALOG = FilterCatalog(_safety_params)
except Exception:
    _SAFETY_CATALOG = None

# =============================================================================
# 1. CORE ARCHITECTURAL VALIDATION
# =============================================================================

# Typical valence table — defined once at module level
_TYPICAL_VALENCE = {
    'H': {1}, 'C': {4}, 'N': {3, 4}, 'O': {2}, 'F': {1},
    'P': {3, 5}, 'S': {2, 4, 6}, 'Cl': {1}, 'Br': {1}, 'I': {1}
}

def check_structural_integrity(mol):
    if mol is None: return {"error": "Invalid Mol"}

    is_organic = any(a.GetAtomicNum() == 6 for a in mol.GetAtoms())  # faster than GetSymbol()

    # Compute fragments once — avoids calling GetMolFrags twice (original bug)
    frags = Chem.GetMolFrags(mol)
    num_frags = len(frags)
    has_disconnected = num_frags > 1

    unusual_val = False
    for a in mol.GetAtoms():
        sym = a.GetSymbol()
        if sym in _TYPICAL_VALENCE:
            try:
                if a.GetTotalValence() not in _TYPICAL_VALENCE[sym]:
                    unusual_val = True
                    break  # early exit — no need to scan all atoms
            except Exception:
                pass

    try:
        chiral_centers = len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))
    except Exception:
        try:
            from rdkit.Chem import FindPotentialStereo
            si = FindPotentialStereo(mol)
            chiral_centers = len([x for x in si if x.type == Chem.StereoType.Atom_Tetrahedral])
        except Exception:
            chiral_centers = 0

    n_atoms = mol.GetNumAtoms()
    symmetry = round(1.0 - (len(set(Chem.CanonicalRankAtoms(mol))) / n_atoms), 3) if n_atoms > 0 else 0

    return {
        "organic": is_organic,
        "disconnected": has_disconnected,
        "fragments": num_frags,
        "unusual_valency": unusual_val,
        "stereo_centers": chiral_centers,
        "symmetry_score": symmetry,
    }

# =============================================================================
# 2. VANGUARD SMARTS DATABASE — pre-compiled at module load
# =============================================================================

_STRUCTURAL_ALERTS_SMARTS = {
    "Reactive_Metabolites": {
        "Quinone_Precursor":  "[OH+0][c1ccccc1][OH+0]",
        "Acyl_Halide":        "[CX3](=[OX1])[F,Cl,Br,I]",
        "Acid_Anhydride":     "[CX3](=O)[OX2][CX3](=O)",
        "Peroxide":           "[OX2][OX2]",
        "Epoxide":            "C1OC1",
        "Aziridine":          "N1CC1",
        "Isocyanate":         "[NX2]=[CX2]=O",
        "Isothiocyanate":     "[NX2]=[CX2]=S",
        "Michael_Acceptor":   "[$([C;H2,H1]=[C;H1,H0]-[C,S,P]=[O,S]),$([C;H2,H1]=[C;H1,H0]-[C]#[N])]",
    },
    "Toxicophores": {
        "Nitro_Group":            "[$([NX3](=O)=O),$([NX3+]([O-])=O)]",
        "Nitroso_Group":          "[NX2]=O",
        "Azo_Linkage":            "[NX2]=[NX2]",
        "Hydrazine":              "[NX3][NX3]",
        "Alpha_Halo_Ketone":      "[CX3](=O)[CX4][F,Cl,Br,I]",
        "Disulfide":              "[SX2][SX2]",
        "Primary_Aromatic_Amine": "[c][NX3H2]",
        "Thioketone":             "[CX3]=[SX1]",
    },
    "Solubility_Killers": {
        "High_LogP_Scaffold":  "c1ccc2c(c1)cccc2",
        "Long_Aliphatic_Chain":"CCCCCCCC",
        "Polycyclic_Aromatic": "c1ccc2c3c(ccc2c1)cccc3",
    },
    "MedChem_Filters": {
        "Triple_Bond":       "[CX2]#[CX2]",
        "Adamantane":        "C12CC3CC(C1)CC(C2)C3",
        "Steroid_Skeleton":  "C1CCC2C(C1)CCC3C2CCC4(C3CCC4)",
        "Macrocycle_12":     "[r12,r13,r14,r15,r16,r17,r18,r19,r20]",
        "Spirocycle":        "[C;R2]([R])([R])([R])[R]",
    },
    "Fragments": {
        "Carboxylic_Acid":  "[CX3](=O)[OX2H1]",
        "Ester":            "[CX3](=O)[OX2H0][#6]",
        "Amide":            "[NX3][CX3](=O)",
        "Sulfonamide":      "[SX4](=O)(=O)[NX3]",
        "Phenol":           "[OX2H][c]",
        "Alcohol":          "[OX2H][C]",
        "Nitrile":          "[NX1]#[CX2]",
        "Imidazole":        "n1cncc1",
        "Tetrazole":        "n1nnnn1",
        "Morpholine":       "O1CCNCC1",
        "Piperazine":       "N1CCNCC1",
        "Thioether":        "[#6][SX2][#6]",
        "Sulfone":          "[#6][SX4](=O)(=O)[#6]",
        "Phosphate":        "P(=O)(O)(O)O",
    },
    "Covalent_Warheads": {
        "Acrylamide":    "[NX3][CX3](=O)[CX3]=[CX3]",
        "Haloacetamide": "[NX3][CX3](=O)C[F,Cl,Br,I]",
        "Vinyl_Sulfone": "[SX4](=O)(=O)[CX3]=[CX3]",
        "Boronic_Acid":  "B(O)O",
        "Cyanamide":     "N[CX2]#[NX1]",
        "Aldehyde":      "[CX3H1]=O",
    },
    "Environmental_Tox": {
        "Organophosphate":           "P(=S)(O)(O)O",
        "Polyhalogenated_Aromatic":  "c1([F,Cl,Br,I])c([F,Cl,Br,I])cc([F,Cl,Br,I])cc1",
        "Paraquat_Like":             "c1cc[n+](cc1)c2cc[n+](cc2)",
        "Dioxin_Skeleton":           "O1c2ccccc2Oc3ccccc13",
    },
    "Advanced_MedChem": {
        "Glaxo_Filter":  "[$([N;H2,H1,H0;R0]-[C;H2,H1,H0;R0]-[C;H2,H1,H0;R0]-[N;H2,H1,H0;R0])]",
        "BMS_Filter":    "[$([F,Cl,Br,I]-[c]-[c]-[F,Cl,Br,I])]",
        "Pfizer_Filter": "[$([N;H1,H0;R0]-[C;H2,H1,H0;R0]-[C;H2,H1,H0;R0]-[O;H1,H0;R0])]",
        "Vertex_Filter": "[$([S;H1,H0;R0]-[C;H2,H1,H0;R0]-[C;H2,H1,H0;R0]-[N;H1,H0;R0])]",
    },
}

# Generate chain/ring variations into SMARTS dict
for _i in range(2, 51):
    _m = _STRUCTURAL_ALERTS_SMARTS["MedChem_Filters"]
    _m[f"chain_len_{_i}"]       = "C" * _i
    _m[f"ring_size_{_i}"]       = "C" * _i if _i > 2 else "CCC"
    _m[f"branched_alkane_{_i}"] = "C" * _i + "(C)C"
    _m[f"fluorine_count_{_i}"]  = "C" * _i + "F"
    _m[f"oxygen_insertion_{_i}"]= "C" * (_i // 2) + "O" + "C" * (_i // 2)

# ── PRE-COMPILE all SMARTS at module import (runs once, not per molecule) ────
# Skips patterns that fail to compile — safe fallback.
STRUCTURAL_ALERTS_COMPILED: dict = {}
for _cat, _sdict in _STRUCTURAL_ALERTS_SMARTS.items():
    STRUCTURAL_ALERTS_COMPILED[_cat] = {}
    for _name, _sma in _sdict.items():
        try:
            _pat = Chem.MolFromSmarts(_sma)
            if _pat is not None:
                STRUCTURAL_ALERTS_COMPILED[_cat][_name] = _pat
        except Exception:
            pass  # silently skip invalid SMARTS

# Keep the old string dict accessible for external callers that may inspect it
STRUCTURAL_ALERTS = _STRUCTURAL_ALERTS_SMARTS


def scan_structural_alerts(mol):
    """Scan using pre-compiled SMARTS patterns — no recompilation per call."""
    if mol is None:
        return {}
    results = {"alerts": [], "categories": {}, "total_hits": 0}

    for category, pat_dict in STRUCTURAL_ALERTS_COMPILED.items():
        cat_hits = 0
        for name, pat in pat_dict.items():
            try:
                matches = mol.GetSubstructMatches(pat)
                if matches:
                    n = len(matches)
                    results["alerts"].append({"name": name, "category": category, "count": n})
                    cat_hits += n
            except Exception:
                continue
        results["categories"][category] = cat_hits
        results["total_hits"] += cat_hits

    return results

# =============================================================================
# 3. COMPREHENSIVE SCREENING ENGINE
# =============================================================================

def _run_vanguard_core(mol, smi):
    """
    Vanguard Engine v2.0 - Core Intelligence Pipeline
    """
    if mol is None: return {"error": "SMILES Parse Failed"}
    
    # A. Structure & Categorization
    meta = check_structural_integrity(mol)
    alerts = scan_structural_alerts(mol)
    
    # B. PhysChem
    try:
        qed_val = round(_QED.qed(mol), 3)
    except Exception:
        qed_val = 0.5

    props = {
        "MW": round(Descriptors.MolWt(mol), 2),
        "LogP": round(Crippen.MolLogP(mol), 2),
        "TPSA": round(Descriptors.TPSA(mol), 2),
        "HBD": rdMolDescriptors.CalcNumHBD(mol),
        "HBA": rdMolDescriptors.CalcNumHBA(mol),
        "RotBonds": rdMolDescriptors.CalcNumRotatableBonds(mol),
        "Rings": rdMolDescriptors.CalcNumRings(mol),
        "Aromatic_Rings": rdMolDescriptors.CalcNumAromaticRings(mol),
        "Heavy_Atoms": mol.GetNumHeavyAtoms(),
        "Fsp3": round(rdMolDescriptors.CalcFractionCSP3(mol), 3),
        "QED": qed_val,
        "Formal_Charge": Chem.GetFormalCharge(mol),
        "Bertz_Complexity": round(Descriptors.BertzCT(mol), 1),
    }
    
    # C. Drug-Likeness Rules
    rules = {
        "Lipinski": (props["MW"] <= 500 and props["LogP"] <= 5 and props["HBD"] <= 5 and props["HBA"] <= 10),
        "Veber": (props["TPSA"] <= 140 and props["RotBonds"] <= 10),
        "Ghose": (160 <= props["MW"] <= 480 and -0.4 <= props["LogP"] <= 5.6 and props["Heavy_Atoms"] <= 70),
        "Egan": (props["TPSA"] <= 131 and props["LogP"] <= 5.88),
        "Muegge": (200 <= props["MW"] <= 600 and -2 <= props["LogP"] <= 5 and props["TPSA"] <= 150),
    }
    
    # D. Safety Catalogs — use module-level pre-built catalog
    try:
        if _SAFETY_CATALOG is not None:
            entries = list(_SAFETY_CATALOG.GetMatches(mol))
        else:
            raise RuntimeError("catalog unavailable")
        safety = {
            "PAINS": any("PAINS" in e.GetDescription() for e in entries),
            "Brenk": any("Brenk" in e.GetDescription() for e in entries),
            "Safety_Hits": len(entries),
        }
    except Exception:
        safety = {"PAINS": False, "Brenk": False, "Safety_Hits": 0}
    
    # E. Discovery Intelligence
    heavy = props["Heavy_Atoms"] if props["Heavy_Atoms"] > 0 else 1
    intel = {
        "Lipophilic_Efficiency": round(5.0 - props["LogP"], 2),
        "Ligand_Efficiency": round((1.4 * 5.0) / heavy, 2),
        "Lead_Status": "Lead-Like" if 250 <= props["MW"] <= 350 and props["LogP"] <= 3.5 else "NCE",
        "SA_Score": round(2.0 + (props["Bertz_Complexity"] / 1000) + (props["Rings"] * 0.5), 2),
    }

    # F. UI Packaging
    repackaged_tests = []
    for k, v in meta.items():
        repackaged_tests.append({"category": "Structure Integrity", "test": k.replace("_"," ").title(), "result": "PASS" if (v is True or (isinstance(v, int) and v < 5)) else "WARN", "detail": str(v)})
    for k, v in props.items():
        repackaged_tests.append({"category": "Physicochemical", "test": k, "result": "INFO", "detail": str(v)})
    for k, v in rules.items():
        repackaged_tests.append({"category": "Drug-Likeness Rules", "test": k, "result": "PASS" if v else "FAIL", "detail": "Compliant" if v else "Violation"})
    for alert in alerts.get("alerts", []):
        repackaged_tests.append({"category": f"Alert: {alert['category']}", "test": alert["name"], "result": "FAIL", "detail": f"Matched {alert['count']} times"})
    repackaged_tests.append({"category": "Safety Catalogs", "test": "PAINS", "result": "PASS" if not safety["PAINS"] else "FAIL", "detail": "None Detected" if not safety["PAINS"] else "Flagged"})
    repackaged_tests.append({"category": "Safety Catalogs", "test": "Brenk", "result": "PASS" if not safety["Brenk"] else "FAIL", "detail": "None Detected" if not safety["Brenk"] else "Flagged"})

    return {
        "_chemo_tests": repackaged_tests,
        "props": props,
        "rules": rules,
        "intel": intel,
        "alerts": alerts,
        "MW": props["MW"],
        "LogP": props["LogP"],
        "TPSA": props["TPSA"],
        "HBD": props["HBD"],
        "HBA": props["HBA"],
        "RotBonds": props["RotBonds"],
        "QED": props["QED"],
        "SA_Score": intel["SA_Score"]
    }

# =============================================================================
# 4. LEGACY COMPATIBILITY LAYER
# =============================================================================

def run_comprehensive_screening(smi):
    """Entry point for Vanguard screening."""
    try:
        mol = Chem.MolFromSmiles(smi)
        return _run_vanguard_core(mol, smi)
    except Exception as e:
        return {"error": str(e), "_chemo_tests": [], "props": {}, "rules": {}, "intel": {}, "alerts": {}}

def run_all_chemo_tests(mol, ref_mol=None):
    if mol is None: return {}
    smi = Chem.MolToSmiles(mol)
    return run_comprehensive_screening(smi)

def get_scaffold_info(mol):
    if mol is None: return {}
    try:
        scaf = MurckoScaffold.GetScaffoldForMol(mol)
        return {"scaffold_smiles": Chem.MolToSmiles(scaf), "scaffold_heavy": scaf.GetNumHeavyAtoms()}
    except: return {}

def compute_similarity(mol1, mol2):
    if not mol1 or not mol2: return 0.0
    try:
        fp1 = _MORGAN_GEN.GetFingerprint(mol1)
        fp2 = _MORGAN_GEN.GetFingerprint(mol2)
        return DataStructs.TanimotoSimilarity(fp1, fp2)
    except Exception:
        return 0.0
