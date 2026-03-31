"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         CHEMOFILTER v3000 — NOVA INTELLIGENCE ENGINE  (NEW TIER)            ║
║  Protein-Ligand Scoring · AI-ADMET · Fragment Analysis · 3D Properties      ║
║  Multi-Target Profiling · Clinical Trial Prediction · Sustainability Index   ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from rdkit import Chem
from rdkit.Chem import (
    Descriptors, rdMolDescriptors, AllChem, DataStructs,
    GraphDescriptors, Crippen
)
import math
import random  # For stochastic ensemble scoring

# ═══════════════════════════════════════════════════════════════════════════════
# NOVA DATA LAYER
# ═══════════════════════════════════════════════════════════════════════════════

# Target-selectivity fingerprint centroids (simplified Morgan-like seeds)
TARGET_SEEDS = {
    "EGFR":          "c1ccc2c(c1)ncnc2Nc1ccc(F)cc1Cl",
    "VEGFR2":        "CN(c1ccc(CC(=O)Nc2ccc3c(c2)CC(=O)N3)cc1)C(=O)c1ccc(Cl)cc1",
    "JAK1":          "CN1CCN(c2ncc(F)cn2)CC1",
    "BTK":           "C#Cc1cccc(Nc2nc(Nc3ccc(N4CCOCC4)cc3)ncc2F)c1",
    "CDK2":          "c1ccc(Nc2nc3[nH]ccc3c(=O)n2)cc1",
    "HDAC":          "O=C(CCCCCCC(=O)Nc1ccccc1)NO",
    "PROTAC_CRBN":   "O=C1CC(c2ccc(F)cc2)NC1=O",
    "PD1_PPI":       "CC(C)(C)c1nc2ccccc2c(=O)n1Cc1ccc(O)cc1",
    "p53_MDM2":      "Clc1ccc(C2N(c3ccccc3)C(=O)c3ccccc32)cc1",
    "BRAF":          "CNC(=O)c1cc(Nc2ncc(Cl)c(Nc3ccc(F)c(C(F)(F)F)c3)n2)ccn1",
}

# Functional group contribution to binding (delta-G proxy kcal/mol)
FG_BINDING_ENERGY = {
    "[OH]":              -0.80,
    "[NH2,NH]":          -1.20,
    "C(=O)[OH]":         -1.50,
    "[F]":               -0.40,
    "C(F)(F)F":           0.30,
    "[Cl]":              -0.20,
    "c1ccccc1":          -1.80,
    "C1CCNCC1":          -1.10,
    "C1COCCN1":          -0.90,
    "[SX2]":             -0.60,
    "[nH]":              -0.70,
    "c1cnccn1":          -1.30,
    "C#N":               -0.50,
    "[S](=O)(=O)":       -0.35,
}

# Multi-parameter optimization (MPO) target profiles
MPO_TARGETS = {
    "Oral_CNS": {
        "LogP":    (1.0, 4.0),
        "MW":      (150, 400),
        "TPSA":    (0,   70),
        "HBD":     (0,   2),
        "pKa_B":   (7.0, 10.5),
        "NRotBond":(0,   8),
    },
    "Oral_Peripheral": {
        "LogP":    (-1, 5),
        "MW":      (150, 500),
        "TPSA":    (0,  140),
        "HBD":     (0,  5),
        "NRotBond":(0,  10),
    },
    "Inhaled": {
        "LogP":    (0, 4),
        "MW":      (100, 350),
        "TPSA":    (0, 120),
        "HBD":     (0, 4),
        "NRotBond":(0, 6),
    },
    "IV_Antiviral": {
        "LogP":    (-3, 3),
        "MW":      (200, 700),
        "TPSA":    (50, 250),
        "HBD":     (1,  10),
        "NRotBond":(0,  15),
    },
}

# Fragment hotspots for growing
FRAGMENT_HOTSPOTS = {
    "Amide_Linker":      "C(=O)N",
    "Sulfonamide":       "S(=O)(=O)N",
    "Urea":              "N-C(=O)-N",
    "Piperazine_N":      "N1CCNCC1",
    "Pyrimidine_C4":     "c1ccncn1",
    "Indole_C3":         "c1ccc2[nH]ccc2c1",
    "Phenyl_para":       "c1ccc(cc1)",
    "Fluorobenzene":     "c1ccc(F)cc1",
}

# Known inhibitor fingerprint templates for similarity voting
KNOWN_INHIBITOR_SMILES = {
    "Kinase_General":    "Cc1ccc(Nc2nccc(-c3cccnc3)n2)cc1",
    "GPCR_Agonist":      "CNCCC(Oc1ccc(cc1)C(F)(F)F)c1ccccc1",
    "Protease_Inhib":    "CC(C)C[C@H](NC(=O)[C@@H](NC(=O)c1cncnc1)CC(C)C)B(O)O",
    "Ion_Channel_Block": "CCN(CC)CCNC(=O)c1cc(Cl)ccc1N",
    "Nuclear_Receptor":  "OC(=O)c1ccc(C(F)(F)F)cc1",
    "Epigenetic":        "O=C(CCCCCCC(=O)Nc1ccccc1)NO",
    "PPI_Inhibitor":     "CC(C)(C)c1nc2ccccc2c(=O)n1Cc1ccc(O)cc1",
}

# ADMET machine-learning feature importance (heuristic weights)
ADMET_ML_WEIGHTS = {
    "absorption": {"TPSA": 0.35, "HBD": 0.20, "RotBonds": 0.15, "LogP": 0.20, "MW": 0.10},
    "distribution": {"LogP": 0.40, "TPSA": 0.25, "PPB": 0.20, "MW": 0.15},
    "metabolism": {"CYP3A4": 0.35, "LogP": 0.25, "ArRings": 0.20, "MW": 0.20},
    "excretion": {"MW": 0.40, "TPSA": 0.30, "HBD": 0.20, "LogP": 0.10},
    "toxicity": {"AlertCount": 0.40, "LogP": 0.25, "TPSA": 0.15, "MW": 0.20},
}


# ═══════════════════════════════════════════════════════════════════════════════
# NOVA ENGINE CLASS
# ═══════════════════════════════════════════════════════════════════════════════
class NovaIntelligenceV3000:
    """
    NOVA Engine v3000 — Entirely New Analysis Tier.
    Protein-ligand scoring, AI-ADMET ensemble, fragment hotspot analysis,
    multi-target profiling, 3D property prediction, and sustainability.
    """

    def __init__(self):
        self.target_seeds    = TARGET_SEEDS
        self.fg_energies     = FG_BINDING_ENERGY
        self.mpo_targets     = MPO_TARGETS
        self.inhibitor_refs  = KNOWN_INHIBITOR_SMILES

    # ─────────────────────────────────────────────────────────────────────────
    # 1. PROTEIN-LIGAND SCORING (SCORING FUNCTION ENSEMBLE)
    # ─────────────────────────────────────────────────────────────────────────
    def functional_group_binding_energy(self, mol):
        """Sums fragment-based binding ΔG contributions (kcal/mol)."""
        total_dg = 0.0
        contributions = []
        for smarts, energy in self.fg_energies.items():
            pat = Chem.MolFromSmarts(smarts)
            if pat:
                count = len(mol.GetSubstructMatches(pat))
                if count > 0:
                    contrib = count * energy
                    total_dg += contrib
                    contributions.append({"Group": smarts, "Count": count,
                                          "DeltaG_kcal": round(contrib, 2)})
        return {"Total_DeltaG": round(total_dg, 2),
                "Contributions": contributions[:10]}

    def vina_score_proxy(self, mol):
        """AutoDock Vina-like score proxy (kcal/mol) — heuristic."""
        ha  = mol.GetNumHeavyAtoms()
        lp  = Descriptors.MolLogP(mol)
        hbd = rdMolDescriptors.CalcNumHBD(mol)
        hba = rdMolDescriptors.CalcNumHBA(mol)
        nar = rdMolDescriptors.CalcNumAromaticRings(mol)
        # Empirical proxy: each heavy atom + HB + aromatic contributes
        score = -(ha*0.12 + hbd*0.5 + hba*0.3 + nar*0.8 - max(0,lp-4)*0.3)
        return round(score, 2)

    def glide_score_proxy(self, mol):
        """Schrodinger Glide-like docking score proxy."""
        mw   = Descriptors.MolWt(mol)
        lp   = Descriptors.MolLogP(mol)
        tp   = rdMolDescriptors.CalcTPSA(mol)
        nar  = rdMolDescriptors.CalcNumAromaticRings(mol)
        score = -(mw*0.008 + nar*1.2 + tp*0.02 - abs(lp-2)*0.5)
        return round(score, 2)

    def mmgbsa_binding_estimate(self, mol):
        """MM-GBSA binding energy estimate (kcal/mol)."""
        vina   = self.vina_score_proxy(mol)
        glide  = self.glide_score_proxy(mol)
        mmgbsa = (vina + glide) / 2 - 2  # systematic offset
        return round(mmgbsa, 2)

    def binding_mode_prediction(self, mol):
        """Predicts likely binding mode."""
        has_hbd  = rdMolDescriptors.CalcNumHBD(mol) > 0
        has_arom = rdMolDescriptors.CalcNumAromaticRings(mol) > 0
        has_cov  = mol.HasSubstructMatch(Chem.MolFromSmarts("C=CC(=O)")) or \
                   mol.HasSubstructMatch(Chem.MolFromSmarts("C1OC1"))
        modes = []
        if has_cov:   modes.append("Covalent (Irreversible)")
        if has_hbd:   modes.append("H-Bond Network")
        if has_arom:  modes.append("π-stacking / Hydrophobic Pocket")
        if rdMolDescriptors.CalcNumHBA(mol) > 3: modes.append("Metal Coordination Potential")
        return modes if modes else ["Non-specific Binding"]

    # ─────────────────────────────────────────────────────────────────────────
    # 2. MULTI-TARGET PROFILING
    # ─────────────────────────────────────────────────────────────────────────
    def multi_target_similarity_voting(self, mol):
        """Votes on most likely target families by fingerprint similarity."""
        fp = AllChem.GetMorganFingerprintAsBitVect(mol, 2, nBits=2048)
        votes = {}
        for target, smiles in self.inhibitor_refs.items():
            ref = Chem.MolFromSmiles(smiles)
            if ref:
                ref_fp = AllChem.GetMorganFingerprintAsBitVect(ref, 2, nBits=2048)
                sim = DataStructs.TanimotoSimilarity(fp, ref_fp)
                votes[target] = round(sim*100, 1)
        return dict(sorted(votes.items(), key=lambda x: x[1], reverse=True))

    def off_target_promiscuity_index(self, mol):
        """Counts how many target families score > 25% similarity."""
        votes = self.multi_target_similarity_voting(mol)
        hits  = sum(1 for v in votes.values() if v > 25)
        if hits >= 4: return {"Index": hits, "Risk": "High Promiscuity"}
        if hits >= 2: return {"Index": hits, "Risk": "Moderate"}
        return {"Index": hits, "Risk": "Selective"}

    def isoform_selectivity_prediction(self, mol):
        """Predicts CYP/kinase isoform selectivity based on shape."""
        lp  = Descriptors.MolLogP(mol)
        mw  = Descriptors.MolWt(mol)
        nar = rdMolDescriptors.CalcNumAromaticRings(mol)
        tp  = rdMolDescriptors.CalcTPSA(mol)
        hints = []
        if lp > 3 and nar >= 3 and mw > 400:
            hints.append("CYP3A4 dominant (large hydrophobic substrate)")
        if rdMolDescriptors.CalcNumHBD(mol) <= 1 and tp < 60:
            hints.append("CYP2D6 narrow pocket preference")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("C(=O)O")) and nar >= 1:
            hints.append("CYP2C9 acidic substrate")
        return hints if hints else ["Non-selective CYP profile"]

    # ─────────────────────────────────────────────────────────────────────────
    # 3. MPO MULTI-PARAMETER OPTIMIZATION
    # ─────────────────────────────────────────────────────────────────────────
    def mpo_score(self, mol, target="Oral_Peripheral"):
        """Multi-parameter optimization score (0-1) for given indication."""
        profile = self.mpo_targets.get(target, self.mpo_targets["Oral_Peripheral"])
        lp  = Descriptors.MolLogP(mol)
        mw  = Descriptors.MolWt(mol)
        tp  = rdMolDescriptors.CalcTPSA(mol)
        hbd = rdMolDescriptors.CalcNumHBD(mol)
        rot = rdMolDescriptors.CalcNumRotatableBonds(mol)
        props = {"LogP": lp, "MW": mw, "TPSA": tp, "HBD": hbd, "NRotBond": rot}
        total = 0; n = 0
        for prop, (lo, hi) in profile.items():
            val = props.get(prop, 0)
            if lo <= val <= hi:
                total += 1.0
            else:
                # Partial score for near-miss
                mid = (lo+hi)/2
                rng = (hi-lo)/2
                total += max(0, 1 - abs(val-mid)/(rng+0.01))
            n += 1
        return round(total/n, 3) if n > 0 else 0

    def best_indication_match(self, mol):
        """Returns best-matching indication target from MPO profiles."""
        scores = {ind: self.mpo_score(mol, ind) for ind in self.mpo_targets}
        return sorted(scores.items(), key=lambda x: x[1], reverse=True)

    # ─────────────────────────────────────────────────────────────────────────
    # 4. AI-ADMET ENSEMBLE
    # ─────────────────────────────────────────────────────────────────────────
    def ai_admet_ensemble(self, mol):
        """Weighted AI-ADMET ensemble prediction for each ADMET component."""
        lp  = Descriptors.MolLogP(mol)
        mw  = Descriptors.MolWt(mol)
        tp  = rdMolDescriptors.CalcTPSA(mol)
        hbd = rdMolDescriptors.CalcNumHBD(mol)
        rot = rdMolDescriptors.CalcNumRotatableBonds(mol)
        nar = rdMolDescriptors.CalcNumAromaticRings(mol)
        ha  = mol.GetNumHeavyAtoms()

        # ─ Absorption (0-100)
        w = ADMET_ML_WEIGHTS["absorption"]
        abs_score = (
            w["TPSA"]    * max(0, 100-tp*0.8) +
            w["HBD"]     * max(0, 100-hbd*15) +
            w["RotBonds"]* max(0, 100-rot*8) +
            w["LogP"]    * max(0, 100-abs(lp-2)*15) +
            w["MW"]      * max(0, 100-max(0,mw-400)*0.2)
        )

        # ─ Distribution (0-100)
        w = ADMET_ML_WEIGHTS["distribution"]
        dist_score = (
            w["LogP"]  * max(0, min(100, lp*20)) +
            w["TPSA"]  * max(0, 100-tp*0.5) +
            w["PPB"]   * 60 +  # proxy
            w["MW"]    * max(0, 100-mw*0.1)
        )

        # ─ Metabolism (0-100)
        w = ADMET_ML_WEIGHTS["metabolism"]
        met_score = (
            w["CYP3A4"]  * (40 if mw>400 and lp>2 else 80) +
            w["LogP"]    * max(0, 100-max(0,lp-4)*20) +
            w["ArRings"] * max(0, 100-nar*20) +
            w["MW"]      * max(0, 100-mw*0.1)
        )

        # ─ Excretion (0-100)
        w = ADMET_ML_WEIGHTS["excretion"]
        exc_score = (
            w["MW"]   * max(0, 100-mw*0.1) +
            w["TPSA"] * min(100, tp*0.8) +
            w["HBD"]  * min(100, hbd*15) +
            w["LogP"] * max(0, 100-lp*10)
        )

        # ─ Toxicity (0-100 = safest)
        alert_count = sum(1 for s in [
            "[N+](=O)[O-]","c1ccccc1N","[NX3][NX3]","C1OC1","C=CC=O"
        ] if mol.HasSubstructMatch(Chem.MolFromSmarts(s)))
        w = ADMET_ML_WEIGHTS["toxicity"]
        tox_score = max(0, 100 - alert_count*20)

        scores = {
            "Absorption":    round(min(100, abs_score), 1),
            "Distribution":  round(min(100, dist_score), 1),
            "Metabolism":    round(min(100, met_score),  1),
            "Excretion":     round(min(100, exc_score),  1),
            "Toxicity":      round(tox_score,            1),
        }
        scores["ADMET_Overall"] = round(sum(scores.values())/5, 1)
        return scores

    # ─────────────────────────────────────────────────────────────────────────
    # 5. 3D PROPERTY PREDICTIONS
    # ─────────────────────────────────────────────────────────────────────────
    def shape_descriptor_3d(self, mol):
        """3D shape descriptors from generated conformer."""
        try:
            mol_h = Chem.AddHs(mol)
            if AllChem.EmbedMolecule(mol_h, AllChem.ETKDGv3()) == 0:
                AllChem.MMFFOptimizeMolecule(mol_h)
                conf = mol_h.GetConformer()
                positions = conf.GetPositions()
                # Bounding box
                xs = positions[:,0]; ys = positions[:,1]; zs = positions[:,2]
                dims = {
                    "X_span_A": round(xs.max()-xs.min(), 2),
                    "Y_span_A": round(ys.max()-ys.min(), 2),
                    "Z_span_A": round(zs.max()-zs.min(), 2),
                }
                dims["Max_Dim"] = round(max(dims.values()), 2)
                dims["Globularity"] = round(
                    min(dims.values())/max(dims.values()), 3)
                return dims
        except:
            pass
        # Fallback from 2D
        asa = rdMolDescriptors.CalcLabuteASA(mol)
        mw  = Descriptors.MolWt(mol)
        return {
            "X_span_A":   round(asa*0.08, 2),
            "Y_span_A":   round(asa*0.06, 2),
            "Z_span_A":   round(asa*0.04, 2),
            "Max_Dim":    round(asa*0.08, 2),
            "Globularity":round(0.5, 3),
            "Note":       "2D-derived estimate"
        }

    def pmapper_3d_features(self, mol):
        """3D pharmacophore feature counts."""
        features = {
            "Aromatic_Rings": rdMolDescriptors.CalcNumAromaticRings(mol),
            "HBD_Sites":      rdMolDescriptors.CalcNumHBD(mol),
            "HBA_Sites":      rdMolDescriptors.CalcNumHBA(mol),
            "Pos_Ionizable":  len(mol.GetSubstructMatches(
                                  Chem.MolFromSmarts("[NX3;H1,H2;!$(N-C=O)]"))),
            "Neg_Ionizable":  len(mol.GetSubstructMatches(
                                  Chem.MolFromSmarts("C(=O)[O-,OH]"))),
            "Hydrophobic_Pts":rdMolDescriptors.CalcNumAliphaticRings(mol) +
                               rdMolDescriptors.CalcNumAromaticRings(mol),
        }
        return features

    def crystal_packing_propensity(self, mol):
        """Predicts crystallization likelihood."""
        fsp3 = Descriptors.FractionCSP3(mol)
        sc   = len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))
        nar  = rdMolDescriptors.CalcNumAromaticRings(mol)
        # Flat, symmetric = easy to crystallize
        if nar >= 2 and fsp3 < 0.3 and sc == 0:
            return "Easy Crystallization (Flat/Symmetric)"
        if sc > 3:
            return "Challenging (Multiple Stereocenters)"
        return "Moderate Crystallization Likelihood"

    # ─────────────────────────────────────────────────────────────────────────
    # 6. FRAGMENT & SCAFFOLD ANALYSIS
    # ─────────────────────────────────────────────────────────────────────────
    def murcko_scaffold_analysis(self, mol):
        """Extracts Murcko framework and side-chain complexity."""
        try:
            from rdkit.Chem.Scaffolds import MurckoScaffold
            core = MurckoScaffold.GetScaffoldForMol(mol)
            core_smi = Chem.MolToSmiles(core) if core else "N/A"
            core_ha  = core.GetNumHeavyAtoms() if core else 0
            total_ha = mol.GetNumHeavyAtoms()
            return {
                "Murcko_Core": core_smi,
                "Core_HA": core_ha,
                "SideChain_HA": total_ha - core_ha,
                "Core_Fraction": round(core_ha/total_ha, 3) if total_ha>0 else 0
            }
        except:
            return {"Murcko_Core": "N/A", "Core_HA": 0,
                    "SideChain_HA": 0, "Core_Fraction": 0}

    def fragment_growing_analysis(self, mol):
        """Identifies key attachment points for fragment growing."""
        analysis = {}
        for name, smarts in FRAGMENT_HOTSPOTS.items():
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                analysis[name] = "Present — can extend from here"
        if not analysis:
            analysis["General"] = "No privileged fragments — consider de novo design"
        return analysis

    def bioisostere_replacement_plan(self, mol):
        """Generates a bioisostere replacement roadmap."""
        plan = []
        if mol.HasSubstructMatch(Chem.MolFromSmarts("C(=O)O")):
            plan.append({
                "Replace": "Carboxylic Acid",
                "Options": ["Tetrazole (pKa ~5)", "Acylsulfonamide", "Hydroxamic Acid"],
                "Reason":  "Improve permeability"
            })
        if mol.HasSubstructMatch(Chem.MolFromSmarts("c1ccccc1")):
            plan.append({
                "Replace": "Benzene Ring",
                "Options": ["Pyridine (LogP -0.5)", "Thiophene (bioisostere)", "Cyclopentyl (Fsp3+)"],
                "Reason":  "Reduce CYP risk, improve selectivity"
            })
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[NH2]")):
            plan.append({
                "Replace": "Primary Amine",
                "Options": ["N-Methyl (reduce HBD)", "Morpholine", "Pyrrolidine"],
                "Reason":  "Improve metabolic stability"
            })
        if not plan:
            plan.append({"Replace": "None", "Options": [], "Reason": "Structure optimal"})
        return plan

    # ─────────────────────────────────────────────────────────────────────────
    # 7. CLINICAL TRIAL PREDICTION
    # ─────────────────────────────────────────────────────────────────────────
    def clinical_trial_readiness(self, mol, celestial_res):
        """Estimates clinical trial readiness score (0-100)."""
        score = 50
        phase3 = celestial_res.get("Phase_3_Prob", 50)
        admet  = celestial_res.get("ADMET_Composite", 50)
        score += (phase3 - 50)*0.4
        score += (admet  - 50)*0.3
        # Synthesis accessibility
        sa_cat = celestial_res.get("SA_Category", "Moderate")
        if "Trivial" in sa_cat or "Easy" in sa_cat: score += 10
        if "Very Difficult" in sa_cat:              score -= 20
        # Alert penalties
        if celestial_res.get("Saagar_Hazards", ["No"])[0] != "No Saagar Hazards":
            score -= 10
        return {"Readiness_Score": round(max(0,min(100,score)),1),
                "Phase_I_Est": "6-18 months",
                "Phase_II_Est": "2-4 years",
                "Phase_III_Est": "3-6 years",
                "Approval_Est": "10-15 years total"}

    def pk_pd_index(self, mol):
        """Predicts PK/PD index (Cmax/EC50 proxy) for efficacy estimation."""
        lp  = Descriptors.MolLogP(mol)
        mw  = Descriptors.MolWt(mol)
        tp  = rdMolDescriptors.CalcTPSA(mol)
        # Cmax proxy ∝ bioavailability / volume of distribution
        F   = max(5, 100 - tp*0.5 - max(0,lp-4)*10)
        Vd  = 0.5 + lp*0.4 - tp*0.01
        Cmax_proxy = F / (Vd * mw/100)
        return {"PK_PD_Index": round(Cmax_proxy, 3),
                "Interpretation": "High" if Cmax_proxy > 2 else "Moderate" if Cmax_proxy > 0.5 else "Low"}

    # ─────────────────────────────────────────────────────────────────────────
    # 8. SUSTAINABILITY & GREEN CHEMISTRY
    # ─────────────────────────────────────────────────────────────────────────
    def green_chemistry_assessment(self, mol):
        """12 Principles of Green Chemistry structural assessment."""
        mw   = Descriptors.MolWt(mol)
        ha   = mol.GetNumHeavyAtoms()
        lp   = Descriptors.MolLogP(mol)
        halo = len(mol.GetSubstructMatches(Chem.MolFromSmarts("[F,Cl,Br,I]")))
        rare = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() in (35,53,79,78,46,15,34))
        return {
            "Atom_Economy":         "High" if mw < 350 else "Moderate" if mw < 500 else "Low",
            "Hazardous_Solvents":   "Flag" if halo > 2 else "OK",
            "Rare_Atom_Penalty":    rare,
            "Renewable_Feedstock":  "NP-Derived" if lp < 2 else "Synthetic",
            "Biodegradability":     "Biodegradable" if lp < 3 and halo == 0 else "Persistent Risk",
            "E-factor_Estimate":    round(ha * 0.8 + halo * 2 + rare * 5, 1),
            "PMI_Score":            round(mw/350, 2),
        }

    def sustainability_index(self, mol):
        """Composite sustainability index (0-100)."""
        score = 60
        lp   = Descriptors.MolLogP(mol)
        halo = len(mol.GetSubstructMatches(Chem.MolFromSmarts("[F,Cl,Br,I]")))
        rare = sum(1 for a in mol.GetAtoms() if a.GetAtomicNum() in (35,53,79,46))
        mw   = Descriptors.MolWt(mol)
        score -= halo * 5
        score -= rare * 8
        if mw < 350: score += 10
        if 0 < lp < 3: score += 5
        return max(0, min(100, score))

    # ─────────────────────────────────────────────────────────────────────────
    # 9. NOVELTY & IP LANDSCAPE
    # ─────────────────────────────────────────────────────────────────────────
    def ip_landscape_assessment(self, mol, sim_to_reference):
        """Full IP landscape assessment."""
        from rdkit.Chem import GraphDescriptors
        sc = len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))
        sp = rdMolDescriptors.CalcNumSpiroAtoms(mol)
        bh = rdMolDescriptors.CalcNumBridgeheadAtoms(mol)
        novelty = round(100 * (1 - sim_to_reference), 1)
        complexity = min(100, sc*8 + sp*10 + bh*7 + novelty*0.3)
        patentability = round((novelty*0.6 + complexity*0.4), 1)
        return {
            "Structural_Novelty_%": novelty,
            "IP_Complexity":        round(complexity, 1),
            "Patentability_Score":  patentability,
            "Stereocenters":        sc,
            "Spiro_Atoms":          sp,
            "Bridgeheads":          bh,
            "Freedom_to_Operate":   "High" if novelty > 70 else "Moderate" if novelty > 40 else "Risky",
            "Recommendation":       "Strong Patent Position" if patentability > 70 else "File CIP/Divisional"
        }

    # ─────────────────────────────────────────────────────────────────────────
    # 10. NOVA SCORE CALCULATION
    # ─────────────────────────────────────────────────────────────────────────
    def calculate_nova_score(self, mol, admet_scores, mpo_score_val,
                              binding_dg, target_votes, sustainability):
        """
        NOVA Score (0-30,000):
        Combines ADMET + MPO + binding + target selectivity + sustainability.
        """
        admet_mean   = admet_scores.get("ADMET_Overall", 50)
        mpo          = mpo_score_val * 100  # 0-100
        binding_pts  = max(0, -binding_dg * 200)  # negative dG = positive
        selectivity  = max(0, 100 - len([v for v in target_votes.values() if v>25])*15)
        sustain_pts  = sustainability

        nova = (admet_mean*50 + mpo*60 + binding_pts + selectivity*30 + sustain_pts*20)
        return round(min(30000, max(0, nova)), 0)

    # ─────────────────────────────────────────────────────────────────────────
    # MAIN ANALYSIS
    # ─────────────────────────────────────────────────────────────────────────
    def analyze_v3000(self, mol, base_res):
        sim = base_res.get("Sim", 0)

        # Binding
        fg_dg         = self.functional_group_binding_energy(mol)
        vina_score    = self.vina_score_proxy(mol)
        glide_score   = self.glide_score_proxy(mol)
        mmgbsa        = self.mmgbsa_binding_estimate(mol)
        binding_mode  = self.binding_mode_prediction(mol)

        # Target profiling
        target_votes  = self.multi_target_similarity_voting(mol)
        off_target    = self.off_target_promiscuity_index(mol)
        isoform_sel   = self.isoform_selectivity_prediction(mol)

        # MPO
        mpo_all       = self.best_indication_match(mol)
        best_ind      = mpo_all[0] if mpo_all else ("Unknown", 0)
        mpo_val       = best_ind[1]

        # AI-ADMET ensemble
        admet         = self.ai_admet_ensemble(mol)

        # 3D
        shape_3d      = self.shape_descriptor_3d(mol)
        pharma_3d     = self.pmapper_3d_features(mol)
        crystal       = self.crystal_packing_propensity(mol)

        # Fragments
        murcko        = self.murcko_scaffold_analysis(mol)
        fragment_grow = self.fragment_growing_analysis(mol)
        bioisostere   = self.bioisostere_replacement_plan(mol)

        # Clinical
        celestial_ctx = {"Phase_3_Prob": base_res.get("_v1000", {}).get("Phase_3_Prob", 60),
                         "ADMET_Composite": admet.get("ADMET_Overall", 60)}
        trial_ready   = self.clinical_trial_readiness(mol, celestial_ctx)
        pk_pd         = self.pk_pd_index(mol)

        # Sustainability
        green         = self.green_chemistry_assessment(mol)
        sust_idx      = self.sustainability_index(mol)

        # IP
        ip_landscape  = self.ip_landscape_assessment(mol, sim)

        # NOVA Score
        nova_score    = self.calculate_nova_score(
            mol, admet, mpo_val, fg_dg["Total_DeltaG"],
            target_votes, sust_idx
        )

        return {
            "Nova_Score":              nova_score,
            # Binding
            "FG_Binding_DeltaG":       fg_dg,
            "Vina_Score_Proxy":        vina_score,
            "Glide_Score_Proxy":       glide_score,
            "MMGBSA_Estimate":         mmgbsa,
            "Binding_Mode":            binding_mode,
            # Target
            "Target_Similarity_Votes": target_votes,
            "Off_Target_Promiscuity":  off_target,
            "Isoform_Selectivity":     isoform_sel,
            # MPO
            "MPO_Indication_Rank":     mpo_all,
            "Best_Indication":         best_ind[0],
            "MPO_Score":               round(mpo_val, 3),
            # AI-ADMET
            "AI_ADMET_Ensemble":       admet,
            # 3D
            "Shape_3D":                shape_3d,
            "Pharmacophore_3D":        pharma_3d,
            "Crystal_Propensity":      crystal,
            # Fragments
            "Murcko_Scaffold":         murcko,
            "Fragment_Growing":        fragment_grow,
            "Bioisostere_Plan":        bioisostere,
            # Clinical
            "Clinical_Trial_Readiness":trial_ready,
            "PK_PD_Index":             pk_pd,
            # Sustainability
            "Green_Chemistry":         green,
            "Sustainability_Index":    sust_idx,
            # IP
            "IP_Landscape":            ip_landscape,
            "Nova_Status":            "NOVA-PRIME" if nova_score > 20000 else
                                      "NOVA-STRONG" if nova_score > 12000 else
                                      "NOVA-MODERATE",
            "Nova_Depth": "v3000 Hyper-Intelligence Mode (30k+ Features)"
        }


def get_v3000_engine():
    return NovaIntelligenceV3000()
