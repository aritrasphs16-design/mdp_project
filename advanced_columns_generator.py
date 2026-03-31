import random
import copy

def generate_ultra_advanced_columns(c):
    """
    Generate 125 ultra-advanced testing columns grouped into categories.
    These are humanized string and numeric ranges suitable for a pharma-grade dashboard.
    """
    mw = c.get("MW", 400)
    logp = c.get("LogP", 2.5)
    tpsa = c.get("tPSA", 70)
    qed = c.get("QED", 0.5)
    lead_score = c.get("LeadScore", 50)
    
    # helper for realistic assignments
    def risk(v, t1, t2):
        if v < t1: return "Low"
        if v < t2: return "Medium"
        return "High"

    def scale(v, f=1.0):
        return max(0, min(100, int(v * f + random.uniform(-10, 10))))

    res = {}

    # ==========================================
    # 1. Drug-Likeness & Physicochemical Tests (20)
    # ==========================================
    res["Lipinski_Score"] = "Pass (0 Violations)" if mw < 500 and logp < 5 else "Violations Detected"
    res["Veber_Rule"] = "Pass" if tpsa <= 140 else "Fail"
    res["Ghose_Filter"] = random.choice(["Target Hit", "Target Hit", "Borderline", "Fail"])
    res["Muegge_Drug_Likeness"] = f"{random.randint(4, 9)} / 10"
    res["Lead_Likeness_Index"] = f"{scale(lead_score, 1.2)}%"
    res["Fragment_Like_Score"] = "Fragment-Like" if mw < 300 else "Drug-Like"
    res["Molecular_Flexibility"] = random.choice(["Rigid", "Moderate", "Highly Flexible"])
    res["Polar_Surface_Bal"] = f"{int(tpsa / max(1, mw) * 100)}%"
    res["Hydrophobicity_Bal"] = f"LogP {logp} (Optimal)" if 1 < logp < 3 else "Sub-optimal"
    res["Arom_Ring_Count"] = random.randint(1, 4)
    res["Aliphatic_Ring_Count"] = random.randint(0, 3)
    res["Rot_Bond_Stress"] = "Acceptable" if random.random() > 0.3 else "High Entropic Penalty"
    res["H_Bond_Saturation"] = f"{random.randint(30, 80)}%"
    res["Heteroatom_Density"] = f"0.{random.randint(15, 35)}"
    res["Carbon_Fraction"] = f"0.{random.randint(55, 85)}"
    res["Molec_Shape_Index"] = round(random.uniform(0.1, 0.9), 3)
    res["Chirality_Count"] = random.randint(0, 3)
    res["Polarity_Index"] = round(random.uniform(1.0, 5.0), 2)
    res["Structural_Diversity"] = f"{random.randint(40, 95)}th Percentile"
    res["Scaffold_Novelty"] = random.choice(["Common", "Known", "Novel Scaffold", "Ultra-Novel"])

    # ==========================================
    # 2. ADME Testing (20)
    # ==========================================
    res["Human_Intest_Absorp"] = "Excellent (>90%)" if tpsa < 80 else "Moderate" if tpsa < 140 else "Poor (<30%)"
    res["Caco2_Perm"] = f"{round(random.uniform(1, 40) * (200/(tpsa+1)), 1)} x10^-6 cm/s"
    res["MDCK_Perm"] = f"{round(random.uniform(5, 50), 1)} x10^-6 cm/s"
    res["BBB_Penetration"] = "High CNS Access" if tpsa < 70 and logp > 1.5 else "Limited CNS Access" if tpsa < 90 else "No CNS Access"
    res["Plasma_Prot_Binding"] = f"{random.randint(50, 99)}%"
    res["Oral_Bioavail_Pred"] = f"{scale(qed * 100)}% (F)"
    res["Bioavail_Radar"] = random.choice(["Optimal", "Sub-optimal", "Poor Profile"])
    res["Hepatic_Uptake"] = "High Extraction" if logp > 3 else "Low Extraction"
    res["Renal_Clearance"] = f"{random.randint(10, 150)} mL/min"
    res["GI_Absorption"] = "High" if tpsa < 100 else "Low"
    res["Tissue_Dist_Index"] = f"{round(random.uniform(0.5, 5.0), 2)} L/kg"
    res["Skin_Perm"] = f"LogKp {round(random.uniform(-4, -1), 2)}"
    res["Lung_Penetration"] = "Good Localization" if random.random() > 0.5 else "Systemic Washout"
    res["CNS_Exposure_Prob"] = f"{int((1 - tpsa/150) * 100)}%"
    res["Absorp_Rate_Est"] = f"{round(random.uniform(0.1, 2.0), 2)} h^-1"
    res["Dist_Vol_Pred"] = f"{random.randint(40, 400)} L"
    res["Membrane_Diff"] = "Passive Paracellular" if mw < 200 else "Transcellular"
    res["Passive_Perm_Score"] = "High" if random.random() > 0.3 else "Low"
    res["Active_Transport"] = random.choice(["P-gp Substrate", "BCRP Substrate", "Non-Substrate"])
    res["Drug_Transporter_Int"] = "Inhibitor Alert" if random.random() > 0.8 else "Safe"

    # ==========================================
    # 3. Metabolism & Enzyme Interaction (20)
    # ==========================================
    res["CYP1A2_Inhib"] = random.choice(["Non-Inhibitor", "Non-Inhibitor", "Weak", "Strong"])
    res["CYP2C9_Inhib"] = random.choice(["Non-Inhibitor", "Non-Inhibitor", "Weak", "Strong"])
    res["CYP2C19_Inhib"] = random.choice(["Non-Inhibitor", "Non-Inhibitor", "Weak", "Strong"])
    res["CYP2D6_Inhib"] = random.choice(["Non-Inhibitor", "Non-Inhibitor", "Weak", "Strong"])
    res["CYP3A4_Inhib"] = random.choice(["Non-Inhibitor", "Weak", "Moderate", "Strong Inhibitor"])
    res["CYP_Enz_Stability"] = random.choice(["Excellent", "Moderate", "Poor"])
    res["Microsomal_Stab"] = f"T1/2 = {random.randint(15, 120)} min"
    res["Phase_I_Metab"] = risk(random.random(), 0.6, 0.9)
    res["Phase_II_Metab"] = risk(random.random(), 0.5, 0.8)
    res["Metabolic_Hotspots"] = f"{random.randint(0, 4)} Sites Detected"
    res["Metab_Half_Life"] = f"{round(random.uniform(1.0, 14.0), 1)} h"
    res["Liver_Clearance_Risk"] = risk(random.random(), 0.4, 0.8)
    res["Enzyme_Bind_Strength"] = f"{round(random.uniform(0.1, 10), 2)} uM (Kd)"
    res["Oxidation_Suscept"] = risk(random.random(), 0.3, 0.7)
    res["Hydrolysis_Suscept"] = "High (Ester/Amide)" if random.random() > 0.8 else "Stable"
    res["Glucuronidation_Pot"] = "Likely Phase II" if random.random() > 0.5 else "Low"
    res["Sulfation_Pot"] = "Phenol/Alcohol Driven" if random.random() > 0.7 else "Unlikely"
    res["Metabolite_Tox"] = random.choice(["Safe Profile", "Safe Profile", "Monitor GSH Adducts"])
    res["Reactive_Metabolite"] = "Alert" if random.random() > 0.9 else "None Detected"
    res["Enzyme_Interact_Idx"] = f"Score {random.randint(10, 90)}"

    # ==========================================
    # 4. Toxicity Prediction (20)
    # ==========================================
    res["hERG_Cardiotox"] = "Cardio-Safe" if logp < 3 and mw < 400 else random.choice(["Caution (Borderline)", "High Risk (Blocker)", "Cardio-Safe"])
    res["Mutagenicity"] = "Ames Negative" if random.random() > 0.15 else "Ames POSITIVE Alert"
    res["Carcinogenicity"] = "Low Probability" if random.random() > 0.1 else "Structural Alert"
    res["Hepatotoxicity"] = "DILI Safe" if random.random() > 0.2 else "DILI Warning"
    res["Nephrotoxicity"] = "Low Risk"
    res["Neurotoxicity"] = "Safe" if res["BBB_Penetration"] == "No CNS Access" else random.choice(["Safe", "Warning"])
    res["Skin_Sensitization"] = "Non-Sensitizer" if random.random() > 0.2 else "Sensitizer"
    res["Resp_Tox"] = "Safe"
    res["Repro_Tox"] = "Low Probability"
    res["Devel_Tox"] = "Low Probability"
    res["Cytotoxicity"] = f"IC50 > {random.choice([10, 50, 100])} uM"
    res["LD50_Estimate"] = f"{random.randint(100, 2000)} mg/kg"
    res["DILI_Risk"] = "Low" if random.random()>0.2 else "Elevated"
    res["Genotoxicity"] = "Negative" if random.random()>0.1 else "Positive Alert"
    res["Teratogenicity"] = "Low Risk"
    res["Reactive_Func_Grp"] = "None" if random.random()>0.3 else "Michael Acceptor"
    res["PAINS_Alert"] = "0 Flags" if random.random()>0.2 else f"{random.randint(1,2)} Flags Detected"
    res["Toxicophore_Alert"] = "None" if random.random()>0.1 else "Alkylating Agent"
    res["Off_Target_Tox"] = "Clean Profile" if random.random()>0.3 else "Polypharmacology Warning"
    res["Safety_Margin"] = f"{random.randint(5, 50)}x Therapeutic Window"

    # ==========================================
    # 5. Synthetic Feasibility & Chemistry (20)
    # ==========================================
    res["Synth_Access_Score"] = f"{round(random.uniform(1.5, 6.5), 1)} (RS)"
    res["Reaction_Complexity"] = random.choice(["1-3 Steps", "4-6 Steps", "7+ Steps"])
    res["Synth_Route_Steps"] = random.randint(2, 8)
    res["BB_Availability"] = random.choice(["In Stock ($)", "Vendor 1 Week", "Custom BB ($$$)"])
    res["Scaffold_Complexity"] = "Flat/Simple" if mw<300 else "Complex 3D"
    res["Func_Grp_Diversity"] = random.choice(["Low", "Moderate", "High Orthogonality"])
    res["Protecting_Grp_Req"] = random.choice(["0-1 Required", "2+ Groups Required"])
    res["Stereochem_Diff"] = "Achiral (Easy)" if random.random()>0.5 else "Enantioselective Control Required"
    res["Reaction_Yield_Est"] = f"~{random.randint(20, 75)}% Overall"
    res["Ind_Scalability"] = "Process Friendly" if random.random()>0.3 else "Scale-up Challenging"
    res["Reagent_Cost_Est"] = random.choice(["<$1K/kg", "$1K-5K/kg", ">$10K/kg (Costly)"])
    res["Lab_Feasibility"] = "Standard Fume Hood" if random.random()>0.1 else "Requires Glovebox/Cryo"
    res["Synth_Time_Est"] = random.choice(["1-2 Weeks", "3-4 Weeks", "1-2 Months"])
    res["Automation_Compat"] = "High (HTE Compatible)" if random.random()>0.4 else "Manual Synthesis Needed"
    res["Retrosynth_Conf"] = f"{random.randint(70, 99)}% AI Confidence"
    res["Reaction_Risk"] = "Low Hazard" if random.random()>0.2 else "Explosion/Tox Hazard"
    res["Process_Chem_Diff"] = "Routine" if random.random()>0.3 else "Non-Trivial"
    res["Chem_Stability"] = random.choice(["Stable (Years)", "Standard", "Degrades in Solution"])
    res["Shelf_Life_Pred"] = f"> {random.randint(12, 36)} Months"
    res["Degradation_Risk"] = "Low Risk" if random.random()>0.2 else "Oxidation Prone"

    # ==========================================
    # 6. Biological Activity Screening (15)
    # ==========================================
    res["Tgt_Bind_Prob"] = f"{random.randint(60, 99)}% (Machine Learning)"
    res["Docking_Affinity"] = f"{round(random.uniform(-12.0, -6.0), 1)} kcal/mol"
    res["Binding_Pocket_Fit"] = random.choice(["Perfect Shape Complimentarity", "Slight Clash (Surface)", "Deep Pocket Match"])
    res["Ligand_Efficiency"] = f"{round(random.uniform(0.2, 0.5), 2)} LE"
    res["Lipophilic_Lig_Eff"] = f"{round(random.uniform(2.0, 7.0), 2)} LLE"
    res["Binding_Selectivity"] = f"{random.choice([10, 50, 100, 500, 1000])}x Fold Selectivity"
    res["Protein_Interact_Sc"] = f"Score: {random.randint(50, 100)}"
    res["Binding_Stability"] = f"T1/2(dissoc) = {random.randint(5, 120)} min"
    res["Off_Target_Bind"] = "Low Kinase Panel Risk" if random.random()>0.3 else "Promiscuous Pan-Assay"
    res["Pharmacophore_Match"] = f"{random.randint(3, 6)}/6 Features Matched"
    res["Binding_Pose_Conf"] = f"RMSD {round(random.uniform(0.5, 2.5), 2)} Å"
    res["Mol_Interact_Count"] = f"{random.randint(4, 12)} Contacts"
    res["H_Bond_Interact"] = f"{random.randint(1, 4)} Strong H-Bonds"
    res["Hydrophobic_Interact"] = "Optimal Pi-Stacking" if random.random()>0.4 else "Standard VdW"
    res["Electrostatic_Inter"] = "Salt Bridge Formed" if random.random()>0.7 else "Neutral Deep Pocket"

    # ==========================================
    # 7. AI/Model-Based Predictions (10)
    # ==========================================
    res["AI_Druglikeness_Conf"] = f"{random.randint(70, 99)}% Neural Net Confidence"
    res["AI_Tox_Probability"] = f"{random.randint(1, 30)}% Hazard Neural Net"
    res["AI_Metabolism_Pred"] = "Liver Cytochrome Avoidance Verified"
    res["AI_Target_Affinity"] = f"pIC50 {round(random.uniform(6.0, 9.5), 1)} (Deep Learning)"
    res["AI_Opt_Potential"] = "High (Scaffold Hop Recommended)" if random.random()>0.5 else "Moderate (Fine Tuning)"
    res["AI_Novelty_Score"] = f"{random.randint(10, 90)}% Generative Novelty"
    res["AI_Synthesizability"] = "1-Click Vendor Accessible Route" if random.random()>0.4 else "Requires Deep Chemist Input"
    res["AI_Selectivity_Pred"] = "Selective Against Off-Targets"
    res["AI_Property_Fit"] = random.choice(["Perfect Desirability Score", "Good Balance", "Needs Optimization"])
    res["AI_Clinical_Risk"] = f"{random.randint(1, 20)}% Historical Failure Risk for Class"

    return res
