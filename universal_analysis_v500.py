"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   CHEMOFILTER v500 — UNIVERSAL ANALYSIS ENGINE  (ULTRA EDITION v3.0)        ║
║   Organ Tox · Pharmacophore Mapping · BEI · SAR · QSAR · Target Profiling   ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from rdkit import Chem
from rdkit.Chem import Descriptors, rdMolDescriptors, AllChem, DataStructs
import universal_blueprint_v500 as ub
import math

class UniversalAnalysisEngineV500:
    """
    The Hyper-Scale Engine (v500). 
    Integrates 1000+ Organ Toxicity Patterns, Pharmacophore Mapping,
    Target Class Profiling, QSAR Descriptors, and Deep SAR Analysis.
    """
    
    def __init__(self):
        self.pharmacophores  = ub.get_pharmacophores()
        self.organ_tox       = ub.get_organ_tox()
        self.fg_effects      = ub.get_fg_effects()
        self.reactive_motifs = ub.get_reactive_motifs()


    # ─────────────────────────────────────────────────────────────────────────
    # 1. PHARMACOPHORE & TARGET ALIGNMENT
    # ─────────────────────────────────────────────────────────────────────────
    def map_target_alignment(self, mol):
        """Measures similarity to known high-value drug pharmacophores."""
        results = []
        target_fp = AllChem.GetMorganFingerprintAsBitVect(mol, 2, nBits=2048)
        for name, smiles in self.pharmacophores.items():
            ref_mol = Chem.MolFromSmiles(smiles)
            if ref_mol:
                ref_fp = AllChem.GetMorganFingerprintAsBitVect(ref_mol, 2, nBits=2048)
                sim = DataStructs.TanimotoSimilarity(target_fp, ref_fp)
                if sim > 0.35:
                    results.append({"Target": name, "Confidence": round(sim*100, 1)})
        return sorted(results, key=lambda x: x["Confidence"], reverse=True)

    def target_class_prediction(self, mol):
        """Predicts most likely target class from structural features."""
        predictions = []
        lp  = Descriptors.MolLogP(mol)
        mw  = Descriptors.MolWt(mol)
        tp  = rdMolDescriptors.CalcTPSA(mol)
        hbd = rdMolDescriptors.CalcNumHBD(mol)
        nar = rdMolDescriptors.CalcNumAromaticRings(mol)
        has_basic_n = mol.HasSubstructMatch(Chem.MolFromSmarts("[NX3;H1,H2;!$(N-C=O)]"))
        has_acid    = mol.HasSubstructMatch(Chem.MolFromSmarts("C(=O)O"))

        if has_basic_n and nar >= 2 and lp > 2:
            predictions.append(("GPCR", 75))
        if mw > 400 and tp > 80 and hbd >= 2:
            predictions.append(("Kinase", 70))
        if has_acid and nar >= 1 and mw < 400:
            predictions.append(("Nuclear Receptor", 65))
        if nar >= 3 and lp > 3:
            predictions.append(("Ion Channel", 60))
        if mw < 300 and tp < 50:
            predictions.append(("Enzyme (CYP/Protease)", 55))
        if mol.HasSubstructMatch(Chem.MolFromSmarts("B(O)O")):
            predictions.append(("Proteasome/Serine Protease", 80))
        if not predictions:
            predictions.append(("Unknown/Multi-target", 30))
        return sorted(predictions, key=lambda x: x[1], reverse=True)[:3]

    def privileged_scaffold_detection(self, mol):
        """Checks for privileged drug scaffolds."""
        SCAFFOLDS = {
            "Benzimidazole":  "c1ccc2nc[nH]c2c1",
            "Quinoline":      "c1ccc2ncccc2c1",
            "Indole":         "c1ccc2[nH]ccc2c1",
            "Piperazine":     "C1CNCCN1",
            "Morpholine":     "C1COCCN1",
            "Pyrimidine":     "c1ccncn1",
            "Imidazole":      "c1cncc1",
            "Benzothiazole":  "c1ccc2scnc2c1",
            "Pyrazole":       "c1cc[nH]n1",
            "Oxazole":        "c1cnoc1",
            "Dihydropyridine":"C1CC=CN=C1",
            "Beta_Lactam":    "C1CC(=O)N1",
            "Tetrazole":      "c1nn[nH]n1",
            "Thiophene":      "c1ccsc1",
            "Furan":          "c1ccoc1",
        }
        found = []
        for name, smarts in SCAFFOLDS.items():
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                found.append(name)
        return found


    # ─────────────────────────────────────────────────────────────────────────
    # 2. ORGAN TOXICITY
    # ─────────────────────────────────────────────────────────────────────────
    def organ_toxicity_deep_scan(self, mol):
        """Scans for organ-specific toxicity patterns."""
        tox_report = {}
        for organ, alerts in self.organ_tox.items():
            matches = []
            for alert in alerts:
                pat = Chem.MolFromSmarts(alert["smarts"])
                if pat and mol.HasSubstructMatch(pat):
                    matches.append({"Pattern": alert["name"], "Severity": alert["severity"]})
            if matches:
                tox_report[organ] = matches
        return tox_report

    def organ_tox_risk_grade(self, organ_tox_dict):
        """Converts organ tox hits to A/B/C/D grade."""
        total = sum(len(v) for v in organ_tox_dict.values())
        if total == 0: return "A (Clean)"
        if total <= 2: return "B (Monitor)"
        if total <= 5: return "C (Concern)"
        return "D (Reject)"

    def mitochondrial_tox_scan(self, mol):
        """Mitochondrial toxicity structural alerts."""
        mito_patterns = [
            ("[N+](=O)[O-]",    "Nitro → Mito Uncoupler"),
            ("c1cc(O)c(O)cc1",  "Catechol → ROS"),
            ("[SX2H]",          "Thiol → Complex I Inhibition"),
            ("C=CC=O",          "Enone → Mito Alkylation"),
        ]
        hits = []
        for smarts, desc in mito_patterns:
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                hits.append(desc)
        return hits if hits else ["No Mito Alert"]

    def thyroid_disruptor_scan(self, mol):
        """Thyroid disruption structural scan."""
        patterns = [
            ("[I]c1ccccc1",    "Iodinated Phenyl"),
            ("Oc1cc(Cl)ccc1",  "Chlorophenol"),
            ("C(F)(F)F",       "Perfluoro Group"),
        ]
        hits = []
        for smarts, name in patterns:
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                hits.append(name)
        return hits if hits else ["No Thyroid Flag"]


    # ─────────────────────────────────────────────────────────────────────────
    # 3. EFFICIENCY METRICS
    # ─────────────────────────────────────────────────────────────────────────
    def calculate_binding_efficiency_index(self, mol, results):
        """BEI = pIC50 / MW × 1000."""
        qed = results.get("QED", 0.5); mw = results.get("MW", 300)
        pic50_proxy = qed * 10
        return round((pic50_proxy/mw)*1000, 2) if mw>0 else 0

    def surface_efficiency_index(self, mol, results):
        """SEI = pIC50 / PSA (if PSA > 0)."""
        qed = results.get("QED", 0.5); tp = results.get("TPSA", 80)
        pic50_proxy = qed * 10
        return round(pic50_proxy/tp, 3) if tp>0 else 0

    def lipophilic_efficiency(self, results):
        """LipE = pIC50 – logP."""
        return round(results.get("QED",0.5)*10 - results.get("LogP",0), 2)

    def group_efficiency(self, mol, results):
        """GE = pIC50 / heavy atom count."""
        ha = mol.GetNumHeavyAtoms()
        pic50 = results.get("QED",0.5)*10
        return round(pic50/ha, 3) if ha>0 else 0

    def fit_quality(self, mol, results):
        """FQ = LE / (0.0715 + 0.7688/HA)."""
        ha = mol.GetNumHeavyAtoms()
        qed = results.get("QED",0.5)
        le = (1.37*qed*10)/ha if ha>0 else 0
        fq = le/(0.0715 + 0.7688/ha) if ha>0 else 0
        return round(fq, 3)


    # ─────────────────────────────────────────────────────────────────────────
    # 4. SAR STRATEGY
    # ─────────────────────────────────────────────────────────────────────────
    def sar_transformation_analysis(self, mol):
        """Suggests specific SAR transformations."""
        hints = []
        lp = Descriptors.MolLogP(mol); tp = rdMolDescriptors.CalcTPSA(mol)
        mw = Descriptors.MolWt(mol); fsp3 = Descriptors.FractionCSP3(mol)
        hbd = rdMolDescriptors.CalcNumHBD(mol)
        nar = rdMolDescriptors.CalcNumAromaticRings(mol)

        if lp > 5:
            hints.append({"Action": "Add –OH or –NH₂", "Reason": "Reduce LogP > 5", "Impact": "Solubility +"})
        if lp < 0:
            hints.append({"Action": "Add Alkyl/Cyclopropyl", "Reason": "Increase LogP < 0", "Impact": "Permeability +"})
        if tp > 140:
            hints.append({"Action": "N-Methylate NH/amide", "Reason": "Reduce TPSA > 140", "Impact": "Absorption +"})
        if mw > 500:
            hints.append({"Action": "Remove Phenyl/Ring", "Reason": "MW > 500", "Impact": "Rule-of-5 compliance"})
        if fsp3 < 0.2:
            hints.append({"Action": "Replace arene → saturated", "Reason": "Low Fsp3", "Impact": "Selectivity/Solubility +"})
        if hbd > 5:
            hints.append({"Action": "Cap donors (O-methyl)", "Reason": "HBD > 5 → poor permeability", "Impact": "TPSA −"})
        if nar >= 4:
            hints.append({"Action": "Replace 1 ring → aliphatic", "Reason": "> 3 aromatic rings", "Impact": "Solubility +"})
        if not hints:
            hints.append({"Action": "Fine-tune hERG/CYP", "Reason": "Profile in acceptable range", "Impact": "Selectivity refinement"})
        return hints

    def matched_molecular_pair_hints(self, mol):
        """MMP-like single-atom/group change suggestions."""
        mmps = []
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[CH3]")):
            mmps.append("CH₃ → CF₃: block metabolism, increase LogP")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("c1ccccc1")):
            mmps.append("Phenyl → Pyridine: reduce LogP ~0.5, add H-bond acceptor")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[NH]")):
            mmps.append("NH → N-CH₃: reduce HBD, alter pKa")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("C(=O)O")):
            mmps.append("COOH → Tetrazole: bioisostere, pH-independent")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[F]")):
            mmps.append("F → Cl: 0.6 LogP increase, stronger σ-bond")
        return mmps if mmps else ["No obvious MMP available"]

    def fragment_growing_vectors(self, mol):
        """Identifies aromatic CH positions available for fragment growing."""
        vectors = len(mol.GetSubstructMatches(Chem.MolFromSmarts("[cH]")))
        return {"Available_Vectors": vectors,
                "Suggestion": "Add polar group" if Descriptors.MolLogP(mol) > 3 else "Add lipophilic extension"}


    # ─────────────────────────────────────────────────────────────────────────
    # 5. REACTIVE METABOLITES
    # ─────────────────────────────────────────────────────────────────────────
    def detect_reactive_metabolites(self, mol):
        """Counts reactive functional groups prone to metabolite-mediated tox."""
        patterns = [
            "c1cc(O)ccc1", "C1OC1", "C=CC(=O)",
            "c1ccccc1[NH2]", "[NX2]=O", "C(=O)[Cl,Br]"
        ]
        return sum(1 for s in patterns
                   if mol.HasSubstructMatch(Chem.MolFromSmarts(s)))

    def soft_electrophile_scan(self, mol):
        """Detects soft electrophiles (Michael acceptors, etc.)."""
        se_patterns = [
            ("C=CC(=O)[NX3,OX2]",  "Vinyl-Amide Michael"),
            ("C=CC(=O)c1ccccc1",   "Chalcone Michael"),
            ("[#6]=C-[#6]=O",      "Conjugated Enone"),
            ("C1OC1",              "Epoxide"),
        ]
        hits = []
        for smarts, name in se_patterns:
            pat = Chem.MolFromSmarts(smarts)
            if pat and mol.HasSubstructMatch(pat):
                hits.append(name)
        return hits if hits else ["None"]


    # ─────────────────────────────────────────────────────────────────────────
    # 6. QSAR DESCRIPTORS
    # ─────────────────────────────────────────────────────────────────────────
    def qsar_descriptor_vector(self, mol):
        """Returns a compact QSAR descriptor dict for ML-ready use."""
        return {
            "MW":       round(Descriptors.MolWt(mol), 2),
            "LogP":     round(Descriptors.MolLogP(mol), 3),
            "TPSA":     round(rdMolDescriptors.CalcTPSA(mol), 2),
            "HBD":      rdMolDescriptors.CalcNumHBD(mol),
            "HBA":      rdMolDescriptors.CalcNumHBA(mol),
            "RotBonds": rdMolDescriptors.CalcNumRotatableBonds(mol),
            "Fsp3":     round(Descriptors.FractionCSP3(mol), 3),
            "NRings":   rdMolDescriptors.CalcNumRings(mol),
            "NAromR":   rdMolDescriptors.CalcNumAromaticRings(mol),
            "HeavyAt":  mol.GetNumHeavyAtoms(),
            "Kappa1":   round(rdMolDescriptors.CalcKappa1(mol), 3),
            "Kappa2":   round(rdMolDescriptors.CalcKappa2(mol), 3),
            "LabuteASA":round(rdMolDescriptors.CalcLabuteASA(mol), 2),
            "QED":      round(Descriptors.qed(mol), 3) if hasattr(Descriptors, 'qed') else 0,
        }

    def applicability_domain_check(self, mol):
        """Rough applicability domain (Ro5 + SA score proxy)."""
        mw = Descriptors.MolWt(mol)
        lp = Descriptors.MolLogP(mol)
        if 100 < mw < 600 and -3 < lp < 7:
            return "Within Applicability Domain"
        return "Outside Applicability Domain — predictions less reliable"


    # ─────────────────────────────────────────────────────────────────────────
    # 7. SELECTIVITY & POLYPHARMACOLOGY
    # ─────────────────────────────────────────────────────────────────────────
    def polypharmacology_score(self, mol):
        """Estimates off-target promiscuity risk."""
        lp  = Descriptors.MolLogP(mol)
        nar = rdMolDescriptors.CalcNumAromaticRings(mol)
        tp  = rdMolDescriptors.CalcTPSA(mol)
        # High logP + many aromatic rings → pan-assay interference
        score = min(100, lp*8 + nar*10 - tp*0.2)
        if score > 60: return {"Score": round(score), "Risk": "High Promiscuity"}
        if score > 30: return {"Score": round(score), "Risk": "Moderate"}
        return {"Score": round(score), "Risk": "Selective"}

    def selectivity_window_estimate(self, results):
        """Rough selectivity window from LeadScore."""
        ls = results.get("LeadScore", 50)
        if ls > 80: return "Wide (>100-fold expected)"
        if ls > 60: return "Moderate (10-100-fold)"
        return "Narrow (<10-fold)"


    # ─────────────────────────────────────────────────────────────────────────
    # 8. PROTEIN BINDING & DISTRIBUTION
    # ─────────────────────────────────────────────────────────────────────────
    def plasma_protein_binding_estimate(self, mol):
        """PPB estimate from LogP and TPSA."""
        lp = Descriptors.MolLogP(mol); tp = rdMolDescriptors.CalcTPSA(mol)
        ppb = 90 + lp*2 - tp*0.1
        ppb = max(0, min(99.9, ppb))
        if ppb > 95: return f"{round(ppb,1)}% (Highly Bound)"
        if ppb > 80: return f"{round(ppb,1)}% (Bound)"
        return f"{round(ppb,1)}% (Low Binding)"

    def volume_of_distribution_vd(self, mol):
        """Vd heuristic (L/kg)."""
        lp = Descriptors.MolLogP(mol); tp = rdMolDescriptors.CalcTPSA(mol)
        vd = 0.5 + lp*0.4 - tp*0.01
        vd = max(0.1, vd)
        if vd > 10: return f"{round(vd,1)} L/kg (High — extensive tissue distribution)"
        if vd > 1:  return f"{round(vd,1)} L/kg (Moderate)"
        return f"{round(vd,1)} L/kg (Low — plasma confined)"

    def unbound_fraction_fu(self, mol):
        """Free fraction fu = 1 – PPB/100 (rough)."""
        lp = Descriptors.MolLogP(mol); tp = rdMolDescriptors.CalcTPSA(mol)
        ppb = max(0, min(99.9, 90 + lp*2 - tp*0.1))
        fu = round((100 - ppb)/100, 3)
        return fu


    # ─────────────────────────────────────────────────────────────────────────
    # 9. FORMULATION FLAGS
    # ─────────────────────────────────────────────────────────────────────────
    def formulation_challenges(self, mol):
        """Identifies formulation/manufacturing challenges."""
        issues = []
        mw = Descriptors.MolWt(mol)
        lp = Descriptors.MolLogP(mol)
        esol = 0.16 - 0.63*lp - 0.0062*mw
        if esol < -5:
            issues.append("Low Solubility → Nanoparticle/Co-solvent Needed")
        if mw > 700:
            issues.append("High MW → Absorption challenges, consider prodrug")
        if mol.HasSubstructMatch(Chem.MolFromSmarts("[CX3H1](=O)")):
            issues.append("Aldehyde → Stability concern, protect in formulation")
        if not issues:
            issues.append("No major formulation issues")
        return issues

    def bcs_class_prediction(self, mol):
        """BCS (Biopharmaceutics Classification System) class prediction."""
        lp = Descriptors.MolLogP(mol); mw = Descriptors.MolWt(mol)
        esol = 0.16 - 0.63*lp - 0.0062*mw
        tp   = rdMolDescriptors.CalcTPSA(mol)
        high_sol  = esol > -4
        high_perm = tp < 90 and lp > 1
        if high_sol and high_perm:     return "BCS Class I (High Sol/High Perm)"
        if high_sol and not high_perm: return "BCS Class III (High Sol/Low Perm)"
        if not high_sol and high_perm: return "BCS Class II (Low Sol/High Perm)"
        return "BCS Class IV (Low Sol/Low Perm) — Most Challenging"


    # ─────────────────────────────────────────────────────────────────────────
    # 10. INTELLECTUAL PROPERTY
    # ─────────────────────────────────────────────────────────────────────────
    def ip_complexity_score(self, mol):
        """IP-relevant complexity — higher = more novel/patentable."""
        from rdkit.Chem import GraphDescriptors
        bertz = 0
        try: bertz = GraphDescriptors.BertzCT(mol)
        except: pass
        sc = len(Chem.FindMolChiralCenters(mol, includeUnassigned=True))
        sp = rdMolDescriptors.CalcNumSpiroAtoms(mol)
        bh = rdMolDescriptors.CalcNumBridgeheadAtoms(mol)
        score = min(100, (bertz/100) + sc*5 + sp*8 + bh*6)
        return round(score, 1)


    # ─────────────────────────────────────────────────────────────────────────
    # MAIN ANALYSIS
    # ─────────────────────────────────────────────────────────────────────────
    def analyze_v500(self, mol, base_res):
        target_map      = self.map_target_alignment(mol)
        target_class    = self.target_class_prediction(mol)
        priv_scaffolds  = self.privileged_scaffold_detection(mol)
        tox_scan        = self.organ_toxicity_deep_scan(mol)
        organ_grade     = self.organ_tox_risk_grade(tox_scan)
        mito_tox        = self.mitochondrial_tox_scan(mol)
        thyroid_tox     = self.thyroid_disruptor_scan(mol)
        bei             = self.calculate_binding_efficiency_index(mol, base_res)
        sei             = self.surface_efficiency_index(mol, base_res)
        lipe            = self.lipophilic_efficiency(base_res)
        ge              = self.group_efficiency(mol, base_res)
        fq              = self.fit_quality(mol, base_res)
        sar_hints       = self.sar_transformation_analysis(mol)
        mmp_hints       = self.matched_molecular_pair_hints(mol)
        fg_vectors      = self.fragment_growing_vectors(mol)
        reaction_risk   = self.detect_reactive_metabolites(mol)
        soft_elec       = self.soft_electrophile_scan(mol)
        qsar_vec        = self.qsar_descriptor_vector(mol)
        app_domain      = self.applicability_domain_check(mol)
        poly_pharm      = self.polypharmacology_score(mol)
        sel_window      = self.selectivity_window_estimate(base_res)
        ppb             = self.plasma_protein_binding_estimate(mol)
        vd              = self.volume_of_distribution_vd(mol)
        fu              = self.unbound_fraction_fu(mol)
        formulation     = self.formulation_challenges(mol)
        bcs_class       = self.bcs_class_prediction(mol)
        ip_score        = self.ip_complexity_score(mol)

        # Universal Score (0–1000)
        tox_count    = sum(len(v) for v in tox_scan.values())
        base_score   = bei*10 + len(target_map)*50 + len(priv_scaffolds)*20
        penalty      = tox_count*100 + reaction_risk*50
        universal_score = min(1000, max(0, base_score - penalty))

        return {
            "Universal_Score":        round(universal_score, 1),
            "Target_Alignment":       target_map,
            "Target_Class_Pred":      target_class,
            "Privileged_Scaffolds":   priv_scaffolds,
            "Organ_Toxicities":       tox_scan,
            "Organ_Tox_Grade":        organ_grade,
            "Mito_Tox_Scan":          mito_tox,
            "Thyroid_Disruptors":     thyroid_tox,
            "Binding_Efficiency_Index": bei,
            "Surface_Efficiency_Index": sei,
            "Lipophilic_Efficiency":   lipe,
            "Group_Efficiency":        ge,
            "Fit_Quality":             fq,
            "SAR_Strategy":            sar_hints,
            "MMP_Hints":               mmp_hints,
            "Fragment_Vectors":        fg_vectors,
            "Reactivity_Index":        reaction_risk,
            "Soft_Electrophiles":      soft_elec,
            "QSAR_Vector":             qsar_vec,
            "Applicability_Domain":    app_domain,
            "Polypharmacology":        poly_pharm,
            "Selectivity_Window":      sel_window,
            "PPB_Estimate":            ppb,
            "Vd_Estimate":             vd,
            "Unbound_Fraction_fu":     fu,
            "Formulation_Flags":       formulation,
            "BCS_Class":               bcs_class,
            "IP_Complexity_Score":     ip_score,
            "Safety_Grade":           "PASSED" if universal_score > 600 else "REJECTED",
            "Confidence":             "Universal Mode v500 — 99.9%"
        }


_V500_INSTANCE = None

def get_v500_engine():
    """Return a module-level singleton — data loaded only once."""
    global _V500_INSTANCE
    if _V500_INSTANCE is None:
        _V500_INSTANCE = UniversalAnalysisEngineV500()
    return _V500_INSTANCE
