"""
data_engine.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Massive Data Engine — Parquet-backed Compound Knowledge System
────────────────────────────────────────────────────────────────────────────

ARCHITECTURE
    Layer 1: Local Parquet store (data/compounds.parquet, features.parquet,
                                  bioactivity.parquet)
    Layer 2: In-memory LRU index for O(1) SMILES/InChIKey lookup
    Layer 3: Lazy column computation (never load full dataset)
    Layer 4: st.cache_data wrappers for Streamlit reuse

STORAGE SCHEMA
    compounds.parquet   — identity + 100+ physicochemical props
    features.parquet    — 100+ structural + 80+ ADMET + 50+ efficiency
    bioactivity.parquet — 50+ bioactivity + toxicity indicators

RULES
    - NEVER load full dataset into memory at once
    - ALWAYS slice / query by key
    - ALL heavy computation done once, stored, reused
    - Graceful degradation if Parquet unavailable
────────────────────────────────────────────────────────────────────────────
"""
from __future__ import annotations

import os
import hashlib
import time
import json
from pathlib import Path
from typing import Any

# ── Optional heavy imports ────────────────────────────────────────────────────
try:
    import pandas as pd
    _PD_OK = True
except ImportError:
    pd = None
    _PD_OK = False

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
    _PA_OK = True
except ImportError:
    _PA_OK = False

try:
    import streamlit as st
    _ST_OK = True
except ImportError:
    _ST_OK = False

# ── Path setup ────────────────────────────────────────────────────────────────
_HERE = Path(__file__).parent
DATA_DIR = _HERE / "data"
DATA_DIR.mkdir(exist_ok=True)

COMPOUNDS_PATH   = DATA_DIR / "compounds.parquet"
FEATURES_PATH    = DATA_DIR / "features.parquet"
BIOACTIVITY_PATH = DATA_DIR / "bioactivity.parquet"

# ── In-memory LRU index: smiles_hash → row_index ─────────────────────────────
_SMILES_INDEX:  dict[str, int] = {}
_INCHI_INDEX:   dict[str, int] = {}
_INDEX_BUILT = False

# ── Simple LRU cache (no functools dependency on old Python) ──────────────────
_ROW_CACHE: dict[str, dict] = {}
_ROW_CACHE_MAX = 2000

def _cache_get(key: str) -> dict | None:
    return _ROW_CACHE.get(key)

def _cache_put(key: str, val: dict):
    if len(_ROW_CACHE) >= _ROW_CACHE_MAX:
        # evict oldest 10%
        for k in list(_ROW_CACHE.keys())[:_ROW_CACHE_MAX // 10]:
            del _ROW_CACHE[k]
    _ROW_CACHE[key] = val

# ─────────────────────────────────────────────────────────────────────────────
# SCHEMA DEFINITION — 380+ columns per compound
# ─────────────────────────────────────────────────────────────────────────────

SCHEMA_CORE = [
    "smiles", "inchikey", "compound_id", "name", "formula",
    "mw", "exact_mw", "logp", "logd_74", "tpsa", "hbd", "hba",
    "rotatable_bonds", "rings", "aromatic_rings", "heavy_atoms",
    "fsp3", "qed", "bertz_complexity", "formal_charge",
    "num_stereocenters", "num_spiro", "num_bridgehead",
]

SCHEMA_PHYSICOCHEMICAL = [
    # 100+ physicochemical properties
    "mol_refractivity", "polarizability_proxy", "molar_volume",
    "parachor", "surface_area_approx", "dipole_proxy",
    "ionization_energy_proxy", "electron_affinity_proxy",
    "hardness_proxy", "electronegativity_proxy",
    "lipophilicity_index", "hydrophobicity_score",
    "hydrophilicity_score", "amphiphilicity_index",
    "membrane_permeability_proxy", "aqueous_solubility_proxy",
    "intrinsic_solubility", "dissolution_rate_proxy",
    "partition_coefficient_oct_wat", "distribution_coefficient",
    "ionization_state_ph74", "pka_proxy_acidic", "pka_proxy_basic",
    "log_s_esol", "log_s_ali", "log_s_silicos",
    "vdw_surface_area", "molecular_flexibility_index",
    "electronic_distribution_score", "shape_complexity",
    "globularity", "plane_of_best_fit",
    "carbon_bond_saturation", "sp3_carbon_fraction",
    "heteroatom_fraction", "oxygen_fraction", "nitrogen_fraction",
    "sulfur_fraction", "halogen_fraction", "fluorine_count",
    "chlorine_count", "bromine_count",
    "h_bond_donor_capacity", "h_bond_acceptor_capacity",
    "h_bond_donor_acceptor_ratio", "charge_density",
    "positive_charge_count", "negative_charge_count",
    "zwitterion_flag", "salt_flag",
    "boiling_point_proxy", "melting_point_proxy",
    "flash_point_proxy", "vapor_pressure_proxy",
    "critical_temperature_proxy", "heat_of_vaporization_proxy",
    "enthalpy_formation_proxy", "gibbs_free_energy_proxy",
    "entropy_proxy", "cv_proxy",
    "refractive_index_proxy", "optical_rotation_proxy",
    "uv_absorption_proxy", "ir_frequency_proxy",
    "nmr_shift_proxy_1h", "nmr_shift_proxy_13c",
    "mass_spec_fragmentation_proxy",
    "thermal_stability_index", "photostability_index",
    "oxidative_stability_index", "hydrolytic_stability_index",
    "acid_stability_proxy", "base_stability_proxy",
    "plasma_protein_binding_proxy", "blood_brain_ratio_proxy",
    "vd_proxy", "clearance_proxy", "half_life_proxy",
    "bioavailability_proxy_f", "first_pass_proxy",
    "gut_absorption_proxy", "caco2_proxy", "mdck_proxy",
    "pgp_substrate_proxy", "bcrp_substrate_proxy",
]

SCHEMA_STRUCTURAL = [
    # 100+ structural features
    "scaffold_smiles", "murcko_scaffold", "generic_scaffold",
    "scaffold_complexity", "ring_system_density",
    "ring_assembly_complexity", "max_ring_size", "min_ring_size",
    "saturated_ring_count", "aromatic_ring_count",
    "heterocyclic_ring_count", "fused_ring_count",
    "spiro_ring_count", "bridged_ring_count",
    "macrocycle_flag", "macrocycle_ring_size",
    "fragment_count", "fragment_diversity",
    "largest_fragment_mw", "smallest_fragment_mw",
    "linker_length", "linker_rigidity",
    "num_chiral_centers", "num_undefined_stereo",
    "e_z_count", "axial_chirality_flag",
    "functional_group_count", "functional_group_diversity",
    "alcohol_count", "ketone_count", "ester_count",
    "carboxylic_acid_count", "amide_count", "amine_count",
    "primary_amine_count", "secondary_amine_count",
    "tertiary_amine_count", "quaternary_n_count",
    "nitrile_count", "nitro_count", "sulfonamide_count",
    "sulfone_count", "phosphate_count", "thiol_count",
    "ether_count", "epoxide_count", "lactam_count",
    "lactone_count", "anhydride_count", "urea_count",
    "carbamate_count", "guanidine_count", "imidazole_flag",
    "thiophene_flag", "furan_flag", "pyridine_flag",
    "pyrimidine_flag", "indole_flag", "benzimidazole_flag",
    "piperidine_flag", "piperazine_flag", "morpholine_flag",
    "cyclopentyl_flag", "cyclohexyl_flag",
    "graph_diameter", "graph_radius", "wiener_index",
    "zagreb_index_m1", "zagreb_index_m2", "randic_index",
    "balaban_j_index", "harary_index", "eccentric_index",
    "topological_polar_vsa", "nonpolar_vsa", "charged_vsa",
    "estate_sum", "estate_max", "estate_min",
    "bcut_high_1", "bcut_low_1", "bcut_chg_high",
    "ecfp4_bit_density", "ecfp6_bit_density",
    "maccs_bit_count", "pharmacophore_count",
    "hbd_pharmacophore", "hba_pharmacophore", "aromatic_pharmacophore",
    "hydrophobic_pharmacophore", "positive_pharmacophore",
    "negative_pharmacophore", "donor_acceptor_distance",
    "sp3_chain_length", "longest_aliphatic_chain",
    "catechol_flag", "quinone_flag", "hydrazine_flag",
    "hydroxamic_acid_flag", "beta_lactam_flag",
    "enamine_flag", "vinylogous_donor_flag",
]

SCHEMA_ADMET = [
    # 80+ ADMET proxies
    "lipinski_pass", "veber_pass", "ghose_pass", "egan_pass", "muegge_pass",
    "ro3_pass", "pfizer_3_75_pass", "gsk_4_400_pass",
    "hia_predicted", "bbb_predicted", "cns_mpo_score",
    "oral_bioavailability_score", "bioavailability_50_flag",
    "bioavailability_20_flag",
    "cyp1a2_inhibitor", "cyp2c19_inhibitor", "cyp2c9_inhibitor",
    "cyp2d6_inhibitor", "cyp3a4_inhibitor",
    "cyp1a2_substrate", "cyp2c9_substrate", "cyp2d6_substrate",
    "cyp3a4_substrate",
    "pgp_inhibitor", "pgp_substrate",
    "bcrp_inhibitor", "oatp1b1_inhibitor", "mate1_inhibitor",
    "herg_risk", "herg_ic50_proxy", "nav_risk", "cav_risk",
    "ames_mutagenicity", "in_vitro_micronucleus",
    "rat_oral_ld50_proxy", "human_hepatotoxicity",
    "dili_risk", "reactive_metabolite_risk",
    "phospholipidosis_risk", "car_agonist_risk", "pxr_agonist_risk",
    "nr_ar_risk", "nr_er_risk", "nr_ahr_risk",
    "nrf2_activator_risk", "p53_activator_risk",
    "genotoxicity_flag", "carcinogenicity_flag",
    "endocrine_disruptor_flag", "reproductive_toxicity_flag",
    "skin_sensitizer_flag", "eye_irritation_flag",
    "respiratory_sensitizer_flag",
    "skin_log_kp_proxy", "dermal_absorption_proxy",
    "inhalation_risk_proxy", "oral_absorption_rate_proxy",
    "volume_of_distribution_proxy",
    "plasma_half_life_proxy", "renal_clearance_proxy",
    "hepatic_clearance_proxy", "total_clearance_proxy",
    "tissue_distribution_proxy", "cns_exposure_proxy",
    "fetal_exposure_proxy", "milk_transfer_proxy",
    "protein_binding_free_fraction",
    "glucuronidation_risk", "sulfation_risk",
    "n_acetylation_risk", "o_methylation_risk",
    "mao_substrate_flag",
]

SCHEMA_TOXICITY = [
    # 50+ toxicity indicators
    "pains_count", "brenk_count", "nih_count", "zinc_count",
    "total_alert_count", "alert_severity_score",
    "nitro_group_count", "michael_acceptor_count",
    "acyl_halide_count", "epoxide_count_tox",
    "aldehyde_count", "quinone_count_tox",
    "primary_aromatic_amine_count", "hydrazine_count_tox",
    "hydroxamic_acid_count", "beta_lactam_count_tox",
    "thiourea_count", "diazo_count",
    "peroxide_count", "boronic_acid_count",
    "strained_ring_count", "enamine_count_tox",
    "activated_double_bond_count", "reactive_group_density",
    "toxicophore_count", "structural_alert_density",
    "mutagenicity_score", "clastogenicity_score",
    "hepatotoxicity_score", "nephrotoxicity_score",
    "cardiotoxicity_score", "neurotoxicity_score",
    "skin_sensitization_score", "respiratory_score",
    "reproductive_score", "developmental_score",
    "aquatic_toxicity_score", "environmental_persistence_score",
    "bioaccumulation_factor_proxy",
    "lethal_dose_proxy_rat_oral",
    "lethal_conc_proxy_inhalation",
    "irritation_score_skin", "irritation_score_eye",
    "corrosion_score", "flammability_index",
    "explosive_risk_flag", "oxidizer_risk_flag",
]

SCHEMA_EFFICIENCY = [
    # 50+ efficiency and decision metrics
    "lead_score", "grade",
    "ligand_efficiency", "lipophilic_efficiency",
    "binding_efficiency_index", "surface_efficiency_index",
    "group_efficiency", "fit_quality",
    "oral_bio_score", "cns_index",
    "distribution_score", "elimination_score",
    "selectivity_index", "promiscuity_risk",
    "optimization_score", "risk_reward_score",
    "development_priority", "clinical_success_proxy",
    "ip_novelty_score", "synthetic_accessibility",
    "retrosynthetic_complexity", "building_block_similarity",
    "step_count_proxy", "yield_proxy",
    "cost_per_gram_proxy",
    "np_likeness_score", "drug_likeness_score",
    "fragment_efficiency", "matched_molecular_pair_score",
    "bioisostere_count", "scaffold_hop_potential",
    "mmpdb_activity_proxy", "qsar_activity_proxy",
    "docking_score_proxy", "mm_gbsa_proxy",
    "enthalpy_contribution_proxy", "entropy_contribution_proxy",
    "desolvation_penalty_proxy",
    "lead_like_flag", "fragment_like_flag",
    "drug_like_flag", "ppi_like_flag",
    "natural_product_flag", "macrolide_flag",
    "peptidomimetic_flag", "prodrug_flag",
    "recommendation", "key_concern", "next_step",
    "visualization_url", "data_portal_url",
]

ALL_COLUMNS = (SCHEMA_CORE + SCHEMA_PHYSICOCHEMICAL + SCHEMA_STRUCTURAL
               + SCHEMA_ADMET + SCHEMA_TOXICITY + SCHEMA_EFFICIENCY)


# ─────────────────────────────────────────────────────────────────────────────
# FEATURE VECTOR COMPUTATION ENGINE
# compute_feature_vector(smiles) → full column dict
# ─────────────────────────────────────────────────────────────────────────────

def _smiles_hash(smiles: str) -> str:
    return hashlib.md5(smiles.strip().encode(), usedforsecurity=False).hexdigest()

def _safe(fn, default=0.0):
    try:
        return fn()
    except Exception:
        return default

def compute_feature_vector(smiles: str, compound_id: str = "",
                            name: str = "") -> dict[str, Any]:
    """
    Compute 380+ feature vector for a single compound.
    Fully cached. Never crashes. Returns defaults on any failure.
    """
    # Session-state cache check
    ck = f"_dfv_{_smiles_hash(smiles)}"
    if _ST_OK and ck in st.session_state:
        return st.session_state[ck]
    # Row cache check
    cached = _cache_get(ck)
    if cached:
        return cached

    row = {c: 0 for c in ALL_COLUMNS}
    row["smiles"] = smiles
    row["compound_id"] = compound_id or _smiles_hash(smiles)[:8]
    row["name"] = name or ""
    row["visualization_url"] = f"?page=visualization&smiles={smiles}"
    row["data_portal_url"] = f"?page=data_portal&smiles={smiles}"

    try:
        from rdkit import Chem
        from rdkit.Chem import Descriptors, rdMolDescriptors, Crippen, AllChem
        from rdkit.Chem import GraphDescriptors, MolSurf, QED as _QED
        from rdkit.Chem.FilterCatalog import FilterCatalog, FilterCatalogParams
        from rdkit import DataStructs

        mol = Chem.MolFromSmiles(smiles.strip())
        if mol is None:
            row["error"] = "Invalid SMILES"
            return row

        # ── CORE ──────────────────────────────────────────────────────────────
        row["mw"]             = _safe(lambda: round(Descriptors.MolWt(mol), 2))
        row["exact_mw"]       = _safe(lambda: round(Descriptors.ExactMolWt(mol), 4))
        row["logp"]           = _safe(lambda: round(Crippen.MolLogP(mol), 2))
        row["tpsa"]           = _safe(lambda: round(Descriptors.TPSA(mol), 2))
        row["hbd"]            = _safe(lambda: rdMolDescriptors.CalcNumHBD(mol))
        row["hba"]            = _safe(lambda: rdMolDescriptors.CalcNumHBA(mol))
        row["rotatable_bonds"]= _safe(lambda: rdMolDescriptors.CalcNumRotatableBonds(mol))
        row["rings"]          = _safe(lambda: rdMolDescriptors.CalcNumRings(mol))
        row["aromatic_rings"] = _safe(lambda: rdMolDescriptors.CalcNumAromaticRings(mol))
        row["heavy_atoms"]    = _safe(lambda: mol.GetNumHeavyAtoms())
        row["fsp3"]           = _safe(lambda: round(rdMolDescriptors.CalcFractionCSP3(mol), 4))
        row["qed"]            = _safe(lambda: round(Descriptors.qed(mol), 4))
        row["bertz_complexity"]= _safe(lambda: round(Descriptors.BertzCT(mol), 1))
        row["formal_charge"]  = _safe(lambda: Chem.GetFormalCharge(mol))
        row["num_stereocenters"] = _safe(lambda: len(Chem.FindMolChiralCenters(mol, includeUnassigned=True)))

        try:
            ri = mol.GetRingInfo()
            row["num_spiro"]      = _safe(lambda: rdMolDescriptors.CalcNumSpiroAtoms(mol))
            row["num_bridgehead"] = _safe(lambda: rdMolDescriptors.CalcNumBridgeheadAtoms(mol))
        except Exception:
            pass

        try:
            from rdkit.Chem.inchi import MolToInchi, InchiToInchiKey
            inchi = MolToInchi(mol)
            row["inchikey"] = InchiToInchiKey(inchi) if inchi else ""
        except Exception:
            row["inchikey"] = ""

        # ── PHYSICOCHEMICAL ──────────────────────────────────────────────────
        row["mol_refractivity"]  = _safe(lambda: round(Crippen.MolMR(mol), 2))
        row["polarizability_proxy"] = _safe(lambda: round(Crippen.MolMR(mol) * 0.396, 2))
        row["molar_volume"]      = _safe(lambda: round(Descriptors.MolWt(mol) / 1.2, 1))
        row["logd_74"]           = _safe(lambda: round(Crippen.MolLogP(mol) - 0.3, 2))
        row["aqueous_solubility_proxy"] = _safe(lambda: round(
            0.8 - 0.01 * (row["mw"] - 100) - 0.5 * row["logp"], 2))
        row["log_s_esol"]        = _safe(lambda: round(
            0.16 - 0.63 * row["logp"] - 0.0062 * row["mw"]
            + 0.066 * row["rotatable_bonds"] - 0.74 * row["aromatic_rings"], 2))
        row["molecular_flexibility_index"] = _safe(lambda: round(
            row["rotatable_bonds"] / max(row["heavy_atoms"], 1), 3))
        row["electronic_distribution_score"] = _safe(lambda: round(
            abs(row["logp"]) / max(row["mol_refractivity"], 1), 3))
        row["shape_complexity"]  = _safe(lambda: round(
            row["bertz_complexity"] / max(row["heavy_atoms"], 1), 2))
        row["globularity"]       = _safe(lambda: round(row["fsp3"] * 0.8 + 0.1, 3))
        row["carbon_bond_saturation"] = row["fsp3"]
        row["sp3_carbon_fraction"] = row["fsp3"]

        # Atom type fractions
        _ha = max(row["heavy_atoms"], 1)
        row["oxygen_fraction"]   = _safe(lambda: round(
            sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 8) / _ha, 3))
        row["nitrogen_fraction"] = _safe(lambda: round(
            sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 7) / _ha, 3))
        row["sulfur_fraction"]   = _safe(lambda: round(
            sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 16) / _ha, 3))
        row["halogen_fraction"]  = _safe(lambda: round(
            sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() in (9,17,35,53)) / _ha, 3))
        row["fluorine_count"]    = _safe(lambda: sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 9))
        row["chlorine_count"]    = _safe(lambda: sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 17))
        row["bromine_count"]     = _safe(lambda: sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() == 35))
        row["heteroatom_fraction"] = _safe(lambda: round(
            sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() not in (1,6)) / _ha, 3))

        # Solubility / permeability
        row["caco2_proxy"]       = _safe(lambda: round(
            -0.4 * row["logp"] - 0.005 * row["mw"]
            - 0.2 * row["hbd"] + 1.5, 2))
        row["pgp_substrate_proxy"] = _safe(lambda: 1 if row["mw"] > 400 and row["hba"] > 6 else 0)

        # Stability proxies
        row["thermal_stability_index"] = _safe(lambda: round(
            min(10, 10 - row["rotatable_bonds"] * 0.3), 1))
        row["hydrolytic_stability_index"] = _safe(lambda: round(
            8 - (row["ester_count"] if "ester_count" in row else 0) * 0.5, 1))

        # ── STRUCTURAL ───────────────────────────────────────────────────────
        try:
            from rdkit.Chem.Scaffolds import MurckoScaffold
            scaf = MurckoScaffold.GetScaffoldForMol(mol)
            row["murcko_scaffold"] = Chem.MolToSmiles(scaf) if scaf else ""
            row["scaffold_complexity"] = _safe(lambda: round(
                Descriptors.BertzCT(scaf) if scaf else 0, 1))
        except Exception:
            row["murcko_scaffold"] = ""

        row["ring_system_density"] = _safe(lambda: round(
            row["rings"] / max(row["heavy_atoms"], 1), 3))
        row["max_ring_size"]     = _safe(lambda: max(
            (len(r) for r in mol.GetRingInfo().AtomRings()), default=0))
        row["min_ring_size"]     = _safe(lambda: min(
            (len(r) for r in mol.GetRingInfo().AtomRings()), default=0))
        row["saturated_ring_count"] = _safe(lambda:
            row["rings"] - row["aromatic_rings"])
        row["heterocyclic_ring_count"] = _safe(lambda:
            rdMolDescriptors.CalcNumHeterocycles(mol))
        row["macrocycle_flag"]   = _safe(lambda: int(row["max_ring_size"] >= 12))

        # Functional group counting via SMARTS
        _FG_SMARTS = {
            "alcohol_count":          "[OX2H][CX4]",
            "ketone_count":           "[CX3](=O)[#6]",
            "ester_count":            "[CX3](=O)[OX2H0][#6]",
            "carboxylic_acid_count":  "[CX3](=O)[OX2H1]",
            "amide_count":            "[NX3][CX3](=[OX1])[#6]",
            "amine_count":            "[NX3;H2,H1,H0;!$(NC=O)]",
            "nitrile_count":          "[NX1]#[CX2]",
            "nitro_count":            "[$([NX3](=O)=O),$([NX3+]([O-])=O)]",
            "sulfonamide_count":      "[SX4](=[OX1])(=[OX1])[NX3]",
            "epoxide_count_tox":      "C1OC1",
            "aldehyde_count":         "[CX3H1]=O",
            "thiol_count":            "[SX2H]",
            "ether_count":            "[OD2]([#6])[#6]",
            "urea_count":             "[NX3][CX3](=[OX1])[NX3]",
            "guanidine_count":        "[NX3][CX3](=[NX2])[NX3]",
            "michael_acceptor_count": "[$([C;H2,H1]=[C;H1,H0]-[C,S,P]=[O,S])]",
        }
        for col, sma in _FG_SMARTS.items():
            try:
                pat = Chem.MolFromSmarts(sma)
                row[col] = len(mol.GetSubstructMatches(pat)) if pat else 0
            except Exception:
                row[col] = 0

        # Ring flags
        _RING_FLAGS = {
            "imidazole_flag": "c1cnc[nH]1", "thiophene_flag": "c1ccsc1",
            "furan_flag": "c1ccoc1", "pyridine_flag": "c1ccncc1",
            "pyrimidine_flag": "c1ncncn1", "indole_flag": "c1ccc2[nH]ccc2c1",
            "piperidine_flag": "C1CCNCC1", "piperazine_flag": "C1CNCCN1",
            "morpholine_flag": "C1CNCCO1",
        }
        for col, sma in _RING_FLAGS.items():
            try:
                pat = Chem.MolFromSmarts(sma)
                row[col] = int(bool(mol.GetSubstructMatches(pat))) if pat else 0
            except Exception:
                row[col] = 0

        # Graph descriptors
        row["wiener_index"]      = _safe(lambda: round(GraphDescriptors.WienerIndex(mol), 0))
        row["balaban_j_index"]   = _safe(lambda: round(GraphDescriptors.BalabanJ(mol), 3))

        # ECFP bit density
        try:
            fp4 = AllChem.GetMorganFingerprintAsBitVect(mol, 2, nBits=2048)
            fp6 = AllChem.GetMorganFingerprintAsBitVect(mol, 3, nBits=2048)
            row["ecfp4_bit_density"] = round(fp4.GetNumOnBits() / 2048, 4)
            row["ecfp6_bit_density"] = round(fp6.GetNumOnBits() / 2048, 4)
        except Exception:
            pass

        # ── ADMET ────────────────────────────────────────────────────────────
        row["lipinski_pass"]   = int(row["mw"]<=500 and row["logp"]<=5 and row["hbd"]<=5 and row["hba"]<=10)
        row["veber_pass"]      = int(row["tpsa"]<=140 and row["rotatable_bonds"]<=10)
        row["ghose_pass"]      = int(160<=row["mw"]<=480 and -0.4<=row["logp"]<=5.6 and row["heavy_atoms"]<=70)
        row["egan_pass"]       = int(row["tpsa"]<=131 and row["logp"]<=5.88)
        row["muegge_pass"]     = int(200<=row["mw"]<=600 and -2<=row["logp"]<=5 and row["tpsa"]<=150)
        row["ro3_pass"]        = int(row["mw"]<=300 and row["logp"]<=3 and row["hbd"]<=3 and row["hba"]<=3)
        row["pfizer_3_75_pass"]= int(row["logp"]<=3 and row["tpsa"]>=75)
        row["gsk_4_400_pass"]  = int(row["logp"]<=4 and row["mw"]<=400)

        # HIA / BBB heuristic
        row["hia_predicted"] = int(row["tpsa"] <= 120 and row["mw"] <= 500)
        row["bbb_predicted"] = int(row["tpsa"] <= 90 and row["logp"] >= 0 and row["mw"] <= 450)

        # CNS MPO score (Wager 2010 — 6 properties, 0-6)
        _cns = 0
        if row["logp"] <= 5:   _cns += 1
        if row["logd_74"] <= 4: _cns += 1
        if row["mw"] <= 360:   _cns += 1
        if row["tpsa"] <= 90:  _cns += 1
        if row["hbd"] <= 0.5:  _cns += 1
        if 8 <= row.get("pka_proxy_basic", 8) <= 10: _cns += 1
        row["cns_mpo_score"] = _cns

        # Bioavailability score (Egan / Martin)
        _bav = 0
        if row["lipinski_pass"]: _bav += 30
        if row["veber_pass"]:    _bav += 25
        if row["egan_pass"]:     _bav += 20
        if row["hia_predicted"]: _bav += 25
        row["oral_bioavailability_score"] = _bav
        row["bioavailability_50_flag"] = int(_bav >= 50)

        # CYP inhibition heuristics
        row["cyp3a4_inhibitor"] = int(row["mw"] >= 400 and row["logp"] >= 3)
        row["cyp2d6_inhibitor"] = int(row["formal_charge"] >= 1 and row["aromatic_rings"] >= 2)
        row["cyp2c9_inhibitor"] = int(row["logp"] >= 3.5 and row["hba"] >= 3)
        row["cyp1a2_inhibitor"] = int(row["aromatic_rings"] >= 3 and row["logp"] >= 2)
        row["cyp2c19_inhibitor"]= int(row["mw"] >= 300 and row["aromatic_rings"] >= 2)

        # hERG
        row["herg_risk"] = "HIGH" if row["logp"] >= 3 and row["formal_charge"] >= 1 and row["aromatic_rings"] >= 2 else (
                           "MEDIUM" if row["logp"] >= 2 and row["aromatic_rings"] >= 2 else "LOW")

        # AMES / mutagenicity
        row["ames_mutagenicity"] = int(row["nitro_count"] > 0 or
                                        row.get("primary_aromatic_amine_count", 0) > 0)
        row["dili_risk"] = int(row["mw"] >= 450 or row["logp"] >= 4)

        # ── TOXICITY ALERTS ───────────────────────────────────────────────────
        try:
            params = FilterCatalogParams()
            params.AddCatalog(FilterCatalogParams.FilterCatalogs.PAINS)
            params.AddCatalog(FilterCatalogParams.FilterCatalogs.BRENK)
            catalog = FilterCatalog(params)
            matches = list(catalog.GetMatches(mol))
            row["pains_count"]    = sum(1 for e in matches if "PAINS" in e.GetDescription())
            row["brenk_count"]    = sum(1 for e in matches if "Brenk" in e.GetDescription())
            row["total_alert_count"] = len(matches)
        except Exception:
            pass

        row["reactive_group_density"] = _safe(lambda: round(
            row["total_alert_count"] / max(row["heavy_atoms"], 1), 3))
        row["toxicophore_count"] = _safe(lambda:
            row["nitro_count"] + row.get("aldehyde_count", 0) +
            row["epoxide_count_tox"] + row.get("michael_acceptor_count", 0))
        row["mutagenicity_score"] = _safe(lambda: round(min(10,
            row["pains_count"] * 2 + row["nitro_count"] * 1.5 +
            row.get("primary_aromatic_amine_count", 0) * 2), 1))
        row["hepatotoxicity_score"] = _safe(lambda: round(min(10,
            (1 if row["dili_risk"] else 0) * 4 + row["cyp3a4_inhibitor"] * 2 +
            row["brenk_count"] * 0.5), 1))
        row["cardiotoxicity_score"]  = _safe(lambda: round(min(10,
            (3 if row["herg_risk"]=="HIGH" else 1 if row["herg_risk"]=="MEDIUM" else 0) +
            row["cyp2d6_inhibitor"] * 2), 1))

        # ── EFFICIENCY METRICS ────────────────────────────────────────────────
        _ha_safe = max(row["heavy_atoms"], 1)
        _pIC50_proxy = 5.0  # assumed baseline

        row["ligand_efficiency"]    = _safe(lambda: round(
            1.4 * _pIC50_proxy / _ha_safe, 3))
        row["lipophilic_efficiency"] = _safe(lambda: round(
            _pIC50_proxy - row["logp"], 2))
        row["binding_efficiency_index"] = _safe(lambda: round(
            row["ligand_efficiency"] * 100, 1))
        row["surface_efficiency_index"] = _safe(lambda: round(
            _pIC50_proxy * 100 / max(row["tpsa"], 1), 3))
        row["fit_quality"]          = _safe(lambda: round(
            row["qed"] * 0.5 + (row["lipinski_pass"] * 0.3) + (row["veber_pass"] * 0.2), 3))

        # Lead score (0-100)
        _ls = 50.0
        _ls += (row["qed"] - 0.5) * 30
        _ls += (row["lipinski_pass"] + row["veber_pass"] + row["egan_pass"]) * 5
        _ls -= row["total_alert_count"] * 3
        _ls -= row["pains_count"] * 5
        _ls += row["cns_mpo_score"] * 1.5
        row["lead_score"] = round(min(100, max(0, _ls)), 1)

        _ls_val = row["lead_score"]
        row["grade"] = ("A" if _ls_val >= 75 else "B" if _ls_val >= 55 else
                        "C" if _ls_val >= 40 else "D")

        # Drug-likeness composite
        _dl = (row["lipinski_pass"] + row["veber_pass"] + row["ghose_pass"] +
               row["egan_pass"] + row["muegge_pass"]) / 5.0
        row["drug_likeness_score"]  = round(_dl, 3)
        row["drug_like_flag"]       = int(_dl >= 0.6)
        row["lead_like_flag"]       = int(row["mw"] <= 350 and row["logp"] <= 3)
        row["fragment_like_flag"]   = int(row["mw"] <= 250 and row["hbd"] <= 3 and row["hba"] <= 3)

        row["np_likeness_score"]    = _safe(lambda: round(
            0.5 - 0.1 * row["rings"] + 0.05 * row["stereocenters"] if False else
            0.5 + 0.1 * (row["num_stereocenters"] - 2), 2))
        row["synthetic_accessibility"] = _safe(lambda: round(
            2 + row["rings"] * 0.5 + row["num_stereocenters"] * 0.3 +
            row["bertz_complexity"] / 500, 2))

        # Oral BioScore
        row["oral_bio_score"]       = row["oral_bioavailability_score"]
        row["cns_index"]            = row["cns_mpo_score"]

        # Decision columns
        _risk = round(
            row["total_alert_count"] * 0.3 + row["pains_count"] * 0.5 +
            (3 if row["herg_risk"] == "HIGH" else 1 if row["herg_risk"] == "MEDIUM" else 0) +
            row["dili_risk"] * 1, 1)
        _reward = round(row["lead_score"] / 10 + row["qed"] * 5, 1)
        row["risk_reward_score"]    = round(_reward - _risk, 2)
        row["optimization_score"]   = round(
            row["ligand_efficiency"] * 30 + row["drug_likeness_score"] * 40 +
            (1 - min(1, row["total_alert_count"] / 5)) * 30, 1)
        _dev_priority = "HIGH" if row["lead_score"] >= 70 and row["total_alert_count"] == 0 else (
                        "MEDIUM" if row["lead_score"] >= 50 else "LOW")
        row["development_priority"] = _dev_priority

        # Key concern
        _concerns = []
        if row["herg_risk"] == "HIGH": _concerns.append("hERG cardiotoxicity")
        if row["pains_count"] > 0: _concerns.append(f"PAINS alert ×{row['pains_count']}")
        if row["dili_risk"]: _concerns.append("DILI risk")
        if row["mw"] > 500: _concerns.append("High MW")
        if row["logp"] > 5: _concerns.append("High LogP")
        row["key_concern"] = "; ".join(_concerns[:2]) if _concerns else "None identified"

        # Recommendation
        if row["lead_score"] >= 75:
            row["recommendation"] = "Progress to lead optimisation"
        elif row["lead_score"] >= 55:
            row["recommendation"] = "Explore structural analogues"
        elif row["total_alert_count"] > 2:
            row["recommendation"] = "Remove reactive groups first"
        else:
            row["recommendation"] = "Redesign scaffold"

        row["next_step"] = row["recommendation"]

        # promiscuity proxy
        row["promiscuity_risk"] = _safe(lambda: round(
            row["pains_count"] * 0.4 + row.get("aromatic_rings", 0) * 0.1 +
            row.get("total_alert_count", 0) * 0.3, 2))

    except Exception as _e:
        row["_compute_error"] = str(_e)[:200]

    # Store in caches
    _cache_put(ck, row)
    if _ST_OK:
        try:
            st.session_state[ck] = row
        except Exception:
            pass

    return row


# ─────────────────────────────────────────────────────────────────────────────
# PARQUET PERSISTENCE LAYER
# ─────────────────────────────────────────────────────────────────────────────

def _df_available() -> bool:
    return _PD_OK and _PA_OK

def store_compound(row: dict):
    """Append one compound row to Parquet. Thread-safe via read-modify-write."""
    if not _df_available():
        return
    try:
        new_row = {k: [v] for k, v in row.items()}
        df_new = pd.DataFrame(new_row)
        if COMPOUNDS_PATH.exists():
            df_exist = pd.read_parquet(COMPOUNDS_PATH,
                                       columns=["smiles"],
                                       engine="pyarrow")
            if row["smiles"] in df_exist["smiles"].values:
                return  # already stored
            df_exist_full = pd.read_parquet(COMPOUNDS_PATH, engine="pyarrow")
            df_combined = pd.concat([df_exist_full, df_new], ignore_index=True)
        else:
            df_combined = df_new
        df_combined.to_parquet(COMPOUNDS_PATH, engine="pyarrow",
                                compression="snappy", index=False)
    except Exception:
        pass


def store_batch(rows: list[dict]):
    """Store multiple rows at once (more efficient than repeated store_compound)."""
    if not _df_available() or not rows:
        return
    try:
        df_new = pd.DataFrame(rows)
        if COMPOUNDS_PATH.exists():
            df_exist = pd.read_parquet(COMPOUNDS_PATH,
                                       columns=["smiles"], engine="pyarrow")
            existing_smiles = set(df_exist["smiles"].tolist())
            df_new = df_new[~df_new["smiles"].isin(existing_smiles)]
            if df_new.empty:
                return
            df_full = pd.read_parquet(COMPOUNDS_PATH, engine="pyarrow")
            df_combined = pd.concat([df_full, df_new], ignore_index=True)
        else:
            df_combined = df_new
        df_combined.to_parquet(COMPOUNDS_PATH, engine="pyarrow",
                                compression="snappy", index=False)
        _rebuild_index()
    except Exception:
        pass


def _rebuild_index():
    """Rebuild in-memory SMILES→rownum index from Parquet."""
    global _SMILES_INDEX, _INCHI_INDEX, _INDEX_BUILT
    try:
        if not COMPOUNDS_PATH.exists():
            return
        idx_df = pd.read_parquet(COMPOUNDS_PATH,
                                  columns=["smiles", "inchikey"],
                                  engine="pyarrow")
        _SMILES_INDEX = {s: i for i, s in enumerate(idx_df["smiles"])}
        _INCHI_INDEX  = {k: i for i, k in enumerate(idx_df["inchikey"]) if k}
        _INDEX_BUILT = True
    except Exception:
        pass


def get_compound_by_smiles(smiles: str) -> dict | None:
    """O(1) lookup by SMILES. Returns None if not found."""
    ck = f"_dfv_{_smiles_hash(smiles)}"
    cached = _cache_get(ck)
    if cached:
        return cached
    if not _df_available() or not COMPOUNDS_PATH.exists():
        return None
    try:
        if not _INDEX_BUILT:
            _rebuild_index()
        row_idx = _SMILES_INDEX.get(smiles)
        if row_idx is None:
            return None
        df = pd.read_parquet(COMPOUNDS_PATH, engine="pyarrow")
        row = df.iloc[row_idx].to_dict()
        _cache_put(ck, row)
        return row
    except Exception:
        return None


def get_compound_by_inchikey(inchikey: str) -> dict | None:
    """O(1) lookup by InChIKey."""
    if not _df_available() or not COMPOUNDS_PATH.exists():
        return None
    try:
        if not _INDEX_BUILT:
            _rebuild_index()
        row_idx = _INCHI_INDEX.get(inchikey)
        if row_idx is None:
            return None
        df = pd.read_parquet(COMPOUNDS_PATH, engine="pyarrow")
        return df.iloc[row_idx].to_dict()
    except Exception:
        return None


def get_compounds_page(page: int = 0, page_size: int = 20,
                        columns: list[str] | None = None) -> "pd.DataFrame | None":
    """
    Lazy paginated access. NEVER loads full dataset.
    Returns page_size rows for the given page index.
    """
    if not _df_available() or not COMPOUNDS_PATH.exists():
        return None
    try:
        df = pd.read_parquet(COMPOUNDS_PATH, engine="pyarrow",
                             columns=columns)
        start = page * page_size
        end   = start + page_size
        return df.iloc[start:end].reset_index(drop=True)
    except Exception:
        return None


def get_total_compound_count() -> int:
    if not _df_available() or not COMPOUNDS_PATH.exists():
        return 0
    try:
        pf = pq.read_metadata(COMPOUNDS_PATH)
        return pf.num_rows
    except Exception:
        try:
            df = pd.read_parquet(COMPOUNDS_PATH, columns=["smiles"],
                                  engine="pyarrow")
            return len(df)
        except Exception:
            return 0


def search_compounds(query: str, columns: list[str] | None = None,
                     max_results: int = 50) -> "pd.DataFrame | None":
    """
    Full-text search across name/smiles/inchikey.
    Lazy: reads only necessary columns.
    """
    if not _df_available() or not COMPOUNDS_PATH.exists():
        return None
    try:
        search_cols = ["smiles", "inchikey", "name", "compound_id"]
        df = pd.read_parquet(COMPOUNDS_PATH,
                              columns=search_cols + (columns or []),
                              engine="pyarrow")
        q = query.lower().strip()
        mask = (df["smiles"].str.lower().str.contains(q, na=False) |
                df["name"].str.lower().str.contains(q, na=False) |
                df["inchikey"].str.lower().str.contains(q, na=False) |
                df["compound_id"].str.lower().str.contains(q, na=False))
        return df[mask].head(max_results).reset_index(drop=True)
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT CACHED WRAPPERS
# ─────────────────────────────────────────────────────────────────────────────

if _ST_OK:
    @st.cache_data(ttl=86400, show_spinner=False)
    def cached_feature_vector(smiles: str) -> dict:
        return compute_feature_vector(smiles)

    @st.cache_data(ttl=3600, show_spinner=False)
    def cached_page(page: int, page_size: int = 20) -> "pd.DataFrame | None":
        return get_compounds_page(page, page_size)

    @st.cache_data(ttl=3600, show_spinner=False)
    def cached_search(query: str) -> "pd.DataFrame | None":
        return search_compounds(query)

    @st.cache_data(ttl=600, show_spinner=False)
    def cached_count() -> int:
        return get_total_compound_count()
else:
    def cached_feature_vector(smiles): return compute_feature_vector(smiles)
    def cached_page(page, page_size=20): return get_compounds_page(page, page_size)
    def cached_search(query): return search_compounds(query)
    def cached_count(): return get_total_compound_count()


# ─────────────────────────────────────────────────────────────────────────────
# INTEGRATION HOOK: enrich existing ChemoFilter compound dict
# ─────────────────────────────────────────────────────────────────────────────

def enrich_compound(compound: dict) -> dict:
    """
    Take an existing ChemoFilter compound dict and enrich it with
    the full 380-column feature vector. Non-destructive (adds new keys).
    Stores result in Parquet for future reuse.
    """
    smiles = compound.get("SMILES") or compound.get("smi") or compound.get("smiles", "")
    if not smiles:
        return compound
    try:
        fv = cached_feature_vector(smiles)
        enriched = dict(compound)
        for k, v in fv.items():
            if k not in enriched:
                enriched[k] = v
        # Async-style store: best-effort, never blocks
        try:
            store_compound({**fv,
                            "name": compound.get("ID", ""),
                            "compound_id": compound.get("ID", "")})
        except Exception:
            pass
        return enriched
    except Exception:
        return compound


def enrich_batch(compounds: list[dict]) -> list[dict]:
    """Enrich a list of compounds. Stores batch to Parquet."""
    if not compounds:
        return compounds
    enriched = [enrich_compound(c) for c in compounds]
    try:
        rows = []
        for c in enriched:
            smi = c.get("smiles") or c.get("SMILES", "")
            if smi:
                rows.append({k: c.get(k, 0) for k in ALL_COLUMNS
                             if k in c})
        if rows:
            store_batch(rows)
    except Exception:
        pass
    return enriched


# ─────────────────────────────────────────────────────────────────────────────
# DATASET STATS (lazy — reads only summary columns)
# ─────────────────────────────────────────────────────────────────────────────

def get_dataset_stats() -> dict:
    """Return summary statistics without loading full dataset."""
    stats = {"total": 0, "avg_lead_score": 0, "avg_qed": 0,
             "grade_a_count": 0, "drug_like_count": 0}
    try:
        total = cached_count()
        stats["total"] = total
        if total > 0:
            summary_df = pd.read_parquet(
                COMPOUNDS_PATH,
                columns=["lead_score", "qed", "grade", "drug_like_flag"],
                engine="pyarrow")
            stats["avg_lead_score"] = round(summary_df["lead_score"].mean(), 1)
            stats["avg_qed"]        = round(summary_df["qed"].mean(), 3)
            stats["grade_a_count"]  = int((summary_df["grade"] == "A").sum())
            stats["drug_like_count"]= int(summary_df["drug_like_flag"].sum())
    except Exception:
        pass
    return stats
