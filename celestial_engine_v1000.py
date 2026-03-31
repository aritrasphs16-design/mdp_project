"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   CHEMOFILTER v1000 — CELESTIAL INTELLIGENCE ENGINE  (ULTRA EDITION v3.0)   ║
║   PBPK · QUED · Saagar Deep Check · Clinical Prediction · SHAP-like XAI     ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from rdkit import Chem
from rdkit.Chem import Descriptors, rdMolDescriptors, AllChem, DataStructs
import celestial_data_v1000 as cd
import math


class CelestialIntelligenceV1000:
    """
    The Ultimate Celestial Engine (v1000).
    Mechanistic PBPK, QUED propensity, Saagar hazard library,
    clinical predictions, ADME-PK simulation, and XAI breakdown.
    """

    def __init__(self):
        self.saagar      = cd.get_saagar()
        self.pbpk_model  = cd.get_pbpk_model()
        self.qued_rules  = cd.get_qued_rules()
        self.extra_tox   = cd.get_extra_tox()

    # ─────────────────────────────────────────────────────────────────────────
    # 1. PBPK — ABSORPTION & CLEARANCE
    # ─────────────────────────────────────────────────────────────────────────
    def absorption_rate_ka(self, mol):
        """Mechanistic absorption rate constant Ka (hr⁻¹)."""
        tp  = Descriptors.TPSA(mol)
        mw  = Descriptors.MolWt(mol)
        hbd = rdMolDescriptors.CalcNumHBD(mol)
        inv = (tp/20.0) + (mw/200.0) + (hbd*1.5)
        return round(1.0/(inv+0.5), 3)

    def hepatic_intrinsic_clearance_clint(self, mol, ox_exp):
        """Hepatic intrinsic clearance CLint (mL/min/kg)."""
        lp = Descriptors.MolLogP(mol)
        clint = (lp*2.5) + (ox_exp*50)
        return round(max(0.1, clint), 1)

    def renal_clearance_estimate(self, mol):
        """Renal clearance estimate based on ionization and MW."""
        mw  = Descriptors.MolWt(mol)
        tp  = rdMolDescriptors.CalcTPSA(mol)
        hbd = rdMolDescriptors.CalcNumHBD(mol)
        # Low MW + hydrophilic = renal excretion favoured
        clr = max(0, 5 - mw/200 + tp/50 + hbd)
        if clr > 5:  return "High Renal CL (>5 mL/min/kg)"
        if clr > 2:  return "Moderate Renal CL"
        return "Low Renal CL — hepatic dominant"

    def biliary_excretion_flag(self, mol):
        """Predicts biliary excretion likelihood (MW > 500 + polar)."""
        mw = Descriptors.MolWt(mol); tp = rdMolDescriptors.CalcTPSA(mol)
        if mw > 500 and tp > 100:
            return "Biliary Excretion Likely (Enterohepatic Recycling Risk)"
        return "Renal/Hepatic Excretion Dominant"

    def absolute_bioavailability_estimate(self, mol):
        """Estimates %F from gut absorption and first-pass."""
        tp  = rdMolDescriptors.CalcTPSA(mol)
        rot = rdMolDescriptors.CalcNumRotatableBonds(mol)
        lp  = Descriptors.MolLogP(mol)
        mw  = Descriptors.MolWt(mol)
        # Gut absorption (Johnson model proxy)
        if tp <= 60 and rot <= 5: ga = 0.9
        elif tp <= 100: ga = 0.7
        elif tp <= 140: ga = 0.4
        else: ga = 0.15
        # First pass ~ lp-driven hepatic extraction
        fp = min(0.8, max(0.05, (lp*0.1 + 0.1)))
        F = ga * (1 - fp)
        return round(F*100, 1)

    def half_life_estimate(self, mol):
        """Plasma half-life estimate (hours)."""
        lp  = Descriptors.MolLogP(mol)
        mw  = Descriptors.MolWt(mol)
        # Higher MW + lower lp = longer t½ heuristic
        t_half = (mw/100)*1.5 - lp*0.5 + 2
        t_half = max(0.5, t_half)
        if t_half < 4:   return f"~{round(t_half,1)} hr (Short)"
        if t_half < 12:  return f"~{round(t_half,1)} hr (Moderate)"
        return f"~{round(t_half,1)} hr (Long)"

    # ─────────────────────────────────────────────────────────────────────────
    # 2. TISSUE PARTITIONING
    # ─────────────────────────────────────────────────────────────────────────
    def tissue_partitioning_kp(self, lp, tp, mw):
        """Calculates tissue Kp coefficients."""
        kps = {}
        for tissue, weights in self.pbpk_model.items():
            val = (lp*weights[0]) + (tp*weights[1]) + (mw*weights[2])/100
            kps[tissue] = round(math.pow(10, val/10), 2)
        return kps

    def tissue_to_plasma_ratio(self, mol):
        """T:P ratios for key compartments."""
        lp = Descriptors.MolLogP(mol); tp = rdMolDescriptors.CalcTPSA(mol)
        ratios = {
            "Brain":   round(max(0.01, 10**(0.6*lp - 0.015*tp - 0.5)), 2),
            "Lung":    round(max(0.01, 10**(0.3*lp - 0.005*tp + 0.2)), 2),
            "Liver":   round(max(0.01, 10**(0.5*lp - 0.01*tp + 0.3)), 2),
            "Muscle":  round(max(0.01, 10**(0.2*lp - 0.008*tp)), 2),
            "Fat":     round(max(0.01, 10**(0.9*lp - 0.003*tp + 0.1)), 2),
        }
        return ratios

    # ─────────────────────────────────────────────────────────────────────────
    # 3. QUED ELECTRONIC PROPENSITY
    # ─────────────────────────────────────────────────────────────────────────
    def quantum_electronic_propensity(self, mol):
        """QUED electrostatic/binding propensity tags."""
        tags = []
        for rule in self.qued_rules:
            pat = Chem.MolFromSmarts(rule["smarts"])
            if pat and mol.HasSubstructMatch(pat):
                tags.append(f"{rule['name']}: {rule['impact']}")
        return tags if tags else ["Standard Electronic Profile"]

    def electrophilicity_index(self, mol):
        """Heuristic electrophilicity from LUMO-proxy features."""
        score = 0
        e_patterns = [
            ("C=CC(=O)",   3, "Michael Acceptor"),
            ("[CX3](=O)Cl", 4, "Acyl Chloride"),
            ("C1OC1",       2, "Epoxide"),
            ("[NX2]=O",     2, "N-oxide/Nitroso"),
            ("[N+](=O)[O-]",1, "Nitro"),
        ]
        labels = []
        for smarts, pts, name in e_patterns:
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                score += pts; labels.append(name)
        if score > 5:   cat = "Highly Electrophilic"
        elif score > 2: cat = "Moderately Electrophilic"
        else:           cat = "Weakly Electrophilic"
        return {"Score": score, "Category": cat, "Alerts": labels}

    def nucleophilicity_sites(self, mol):
        """Identifies nucleophilic reaction sites."""
        sites = []
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[SX2H]")):    sites.append("Thiol (strong Nu)")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[NX3;H2]")):  sites.append("Primary Amine")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[OX2H]")):    sites.append("Hydroxyl")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[nH]")):      sites.append("Aromatic NH (Pyrrole-type)")
        return sites if sites else ["No Strong Nucleophilic Sites"]

    # ─────────────────────────────────────────────────────────────────────────
    # 4. SAAGAR HAZARD DEEP CHECK
    # ─────────────────────────────────────────────────────────────────────────
    def saagar_hazard_deep_check(self, mol):
        """Deep scan using the Saagar advanced moiety library."""
        hazards = []
        for s in self.saagar:
            pat = Chem.MolFromSmarts(s["smarts"])
            if pat and mol.HasSubstructMatch(pat):
                hazards.append(f"{s['name']}: {s.get('hazard', s.get('property',''))}")
        return hazards if hazards else ["No Saagar Hazards"]

    def extra_tox_deep_scan(self, mol):
        """Extended toxicity pattern scan."""
        hits = []
        for t in self.extra_tox:
            pat = Chem.MolFromSmarts(t["smarts"])
            if pat and mol.HasSubstructMatch(pat):
                hits.append({"Alert": t["name"], "Risk": t["risk"]})
        return hits if hits else [{"Alert": "None", "Risk": "Clean"}]

    # ─────────────────────────────────────────────────────────────────────────
    # 5. CLINICAL PREDICTIONS
    # ─────────────────────────────────────────────────────────────────────────
    def therapeutic_index_ti_proxy(self, v500_score, tox_hits):
        """Heuristic TI proxy."""
        ti = (v500_score/(tox_hits+1))/100
        return round(ti, 2)

    def phase_3_passage_prob(self, mol, res):
        """Composite Phase III passage probability (%)."""
        prob = 70.0
        fsp3 = Descriptors.FractionCSP3(mol)
        mw   = Descriptors.MolWt(mol)
        lp   = res.get("LogP", 0)
        if fsp3 < 0.25: prob -= 15
        if lp > 4:      prob -= 10
        if res.get("SA_Score", 0) > 4: prob -= 10
        if mw < 400:    prob += 5
        _v500 = res.get("_v500", {})
        if _v500.get("Universal_Score", 0) > 800: prob += 20
        if _v500.get("Organ_Tox_Grade", "D").startswith("A"): prob += 10
        return round(max(5.0, min(99.0, prob)), 1)

    def clinical_candidate_grade(self, phase3_prob, v500_score):
        """Overall clinical candidate grade (A-F)."""
        combined = phase3_prob*0.5 + v500_score*0.05
        if combined > 80: return "A — Excellent Clinical Candidate"
        if combined > 65: return "B — Good Candidate"
        if combined > 50: return "C — Moderate — Further Optimization Needed"
        if combined > 35: return "D — Weak — Major Issues"
        return "F — Reject"

    def human_dose_estimate(self, mol):
        """Rough human efficacious dose estimate heuristic (mg)."""
        mw  = Descriptors.MolWt(mol)
        lp  = Descriptors.MolLogP(mol)
        # Higher MW and lower permeability → higher dose needed
        dose = (mw/300)*50 * (1 + max(0, 4-lp)*0.2)
        return round(min(1000, max(1, dose)), 1)

    def pdt_photosensitivity_alert(self, mol):
        """Photo-dynamic therapy / photosensitivity risk."""
        phototox_patterns = [
            ("c1ccc2c(c1)[nH]c1ccccc12",    "Carbazole-PDT"),
            ("c1ccc2ccoc2c1",               "Coumarin"),
            ("c1ccc2c(c1)ccc1ccccc12",      "Anthracene"),
            ("c1ccc2c(c1)oc3ccccc3c2=O",    "Xanthone"),
        ]
        hits = []
        for smarts, name in phototox_patterns:
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                hits.append(name)
        return hits if hits else ["No Phototox Alert"]

    # ─────────────────────────────────────────────────────────────────────────
    # 6. EXPLAINABLE AI (SHAP-like)
    # ─────────────────────────────────────────────────────────────────────────
    def explainable_ai_shap(self, res):
        """SHAP-like breakdown of primary score drivers."""
        drivers = []
        lp = res.get("LogP", 0); mw = res.get("MW", 0); tp = res.get("TPSA", 0)
        qed = res.get("QED", 0.5); fsp3 = res.get("Fsp3", 0.3)

        drivers.append({"Feature": "Lipophilicity (LogP)",
                         "Value": lp,
                         "Impact": "-25%" if lp>3 else "+15%",
                         "Dir": "↓" if lp>3 else "↑"})
        drivers.append({"Feature": "Molecular Weight",
                         "Value": mw,
                         "Impact": "+20%" if mw<400 else "-15%",
                         "Dir": "↑" if mw<400 else "↓"})
        drivers.append({"Feature": "Polar Surface Area",
                         "Value": tp,
                         "Impact": "+10%" if tp<70 else "-10%",
                         "Dir": "↑" if tp<70 else "↓"})
        drivers.append({"Feature": "QED Drug-Likeness",
                         "Value": round(qed,3),
                         "Impact": f"+{round(qed*30)}%",
                         "Dir": "↑"})
        drivers.append({"Feature": "Fsp3 Saturation",
                         "Value": round(fsp3,2),
                         "Impact": "+12%" if fsp3>0.4 else "-8%",
                         "Dir": "↑" if fsp3>0.4 else "↓"})
        return drivers

    def confidence_interval_estimate(self, phase3_prob):
        """95% CI estimate for clinical probability."""
        ci_low  = max(0,  phase3_prob - 12)
        ci_high = min(100, phase3_prob + 12)
        return {"Point_Estimate": phase3_prob,
                "CI_95": f"[{ci_low}%, {ci_high}%]",
                "Reliability": "Heuristic Model (Indicative Only)"}

    # ─────────────────────────────────────────────────────────────────────────
    # 7. ADVANCED ADMET COMPOSITE
    # ─────────────────────────────────────────────────────────────────────────
    def admet_composite_score(self, mol, res):
        """Composite ADMET score (0–100)."""
        score = 50.0
        # Absorption
        tp = res.get("TPSA", 80)
        if tp < 60:   score += 10
        elif tp > 140: score -= 15
        # Distribution
        lp = res.get("LogP", 2)
        if 1 < lp < 4: score += 8
        elif lp > 5:   score -= 10
        # Metabolism
        has_aniline = mol.HasSubstructMatch(Chem.MolFromSmarts("[NH2]-c1ccccc1"))
        if has_aniline: score -= 12
        # Excretion
        mw = res.get("MW", 300)
        if mw < 400: score += 5
        # Toxicity
        v500 = res.get("_v500", {})
        tox  = sum(len(v) for v in v500.get("Organ_Toxicities", {}).values())
        score -= tox*5
        return round(max(0, min(100, score)), 1)

    def compare_to_approved_drugs(self, mol, res):
        """Places compound in context of approved drug space."""
        lp  = res.get("LogP", 2)
        mw  = res.get("MW", 300)
        tp  = res.get("TPSA", 80)
        qed = res.get("QED", 0.5)
        # Median values from approved small molecules
        assessments = []
        if abs(mw - 370) < 100:    assessments.append("MW in approved drug median range (300-450 Da)")
        if abs(lp - 2.5) < 1.5:   assessments.append("LogP within optimal drug-like range")
        if tp < 90:                assessments.append("TPSA consistent with oral drugs")
        if qed > 0.6:              assessments.append("QED above median for approved drugs")
        return assessments if assessments else ["Outside typical approved drug space on key metrics"]

    # ─────────────────────────────────────────────────────────────────────────
    # MAIN ANALYSIS
    # ─────────────────────────────────────────────────────────────────────────
    def analyze_v1000(self, mol, res):
        ka          = self.absorption_rate_ka(mol)
        ox_exp      = res.get("_v50", {}).get("Oxidative_Exposure", 0.5)
        clint       = self.hepatic_intrinsic_clearance_clint(mol, ox_exp)
        renal_cl    = self.renal_clearance_estimate(mol)
        biliary     = self.biliary_excretion_flag(mol)
        bioavail_F  = self.absolute_bioavailability_estimate(mol)
        t_half      = self.half_life_estimate(mol)
        qued_tags   = self.quantum_electronic_propensity(mol)
        electro     = self.electrophilicity_index(mol)
        nucleo      = self.nucleophilicity_sites(mol)
        kp_map      = self.tissue_partitioning_kp(res.get("LogP",2.0), res.get("TPSA",80.0), res.get("MW",300.0))
        tp_ratios   = self.tissue_to_plasma_ratio(mol)
        saagar_hz   = self.saagar_hazard_deep_check(mol)
        extra_tox   = self.extra_tox_deep_scan(mol)
        univ_score  = res.get("_v500", {}).get("Universal_Score", 500)
        tox_count   = len(res.get("_v500", {}).get("Organ_Toxicities", {}))
        ti_proxy    = self.therapeutic_index_ti_proxy(univ_score, tox_count)
        success_prob= self.phase_3_passage_prob(mol, res)
        clin_grade  = self.clinical_candidate_grade(success_prob, univ_score)
        human_dose  = self.human_dose_estimate(mol)
        phototox    = self.pdt_photosensitivity_alert(mol)
        shap_data   = self.explainable_ai_shap(res)
        ci          = self.confidence_interval_estimate(success_prob)
        admet_score = self.admet_composite_score(mol, res)
        drug_context= self.compare_to_approved_drugs(mol, res)

        celestial_score = (univ_score*4) + (success_prob*10) + (ka*200) - (clint*2)

        return {
            "Celestial_Score":        round(celestial_score, 1),
            # PBPK
            "PBPK_Ka":                ka,
            "PBPK_CLint":             clint,
            "Renal_CL":               renal_cl,
            "Biliary_Excretion":      biliary,
            "Bioavailability_F":      f"{bioavail_F}%",
            "Half_Life_Est":          t_half,
            # Electronic
            "QUED_Tags":              qued_tags,
            "Electrophilicity":       electro,
            "Nucleophilic_Sites":     nucleo,
            # Distribution
            "Kp_Ensemble":            kp_map,
            "Tissue_Plasma_Ratios":   tp_ratios,
            # Hazards
            "Saagar_Hazards":         saagar_hz,
            "Extra_Tox_Scan":         extra_tox,
            "Phototox_Alert":         phototox,
            # Clinical
            "Therapeutic_Index":      ti_proxy,
            "Phase_3_Prob":           success_prob,
            "Clinical_Grade":         clin_grade,
            "Human_Dose_Est_mg":      human_dose,
            "CI_95":                  ci,
            "ADMET_Composite":        admet_score,
            "Drug_Context":           drug_context,
            # XAI
            "SHAP_Breakdown":         shap_data,
            "Status":                "STABLE" if success_prob > 60 else "VOLATILE"
        }


_V1000_INSTANCE = None

def get_v1000_engine():
    """Return a module-level singleton — data loaded only once."""
    global _V1000_INSTANCE
    if _V1000_INSTANCE is None:
        _V1000_INSTANCE = CelestialIntelligenceV1000()
    return _V1000_INSTANCE
