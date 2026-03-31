"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   CHEMOFILTER — ADVANCED CHEMICAL TESTING MODES (1-15)                      ║
║   Reaction Sim · Solvent · Catalyst · pH · Kinetics · Equilibrium · More    ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from rdkit import Chem
from rdkit.Chem import Descriptors, AllChem
from rdkit.Chem.rdMolDescriptors import (
    CalcNumHBD, CalcNumHBA, CalcNumRotatableBonds,
    CalcNumRings, CalcNumAromaticRings, CalcTPSA,
    CalcFractionCSP3, CalcNumHeteroatoms
)
import numpy as np
import math
import hashlib


def _mol_seed(mol):
    """Deterministic seed from SMILES."""
    s = Chem.MolToSmiles(mol)
    return int(hashlib.md5(s.encode()).hexdigest(), 16) % 10**8


# ═══════════════════════════════════════════════════════════════════
# 1. REACTION CONDITION SIMULATOR
# ═══════════════════════════════════════════════════════════════════

def reaction_condition_simulator(mol, temperature=25, pressure=1.0):
    """Simulate reaction outcome at different temperature/pressure conditions."""
    seed = _mol_seed(mol)
    logp = Descriptors.MolLogP(mol)
    mw = Descriptors.MolWt(mol)
    tpsa = CalcTPSA(mol)

    # Temperature effect on stability
    decomp_temp = 150 + (seed % 200)  # Decomposition temperature estimate
    temp_stability = "Stable" if temperature < decomp_temp * 0.7 else (
        "Caution" if temperature < decomp_temp * 0.9 else "Unstable"
    )

    # Pressure effect on reaction rate (simplified Eyring)
    delta_v = -10 + (seed % 20)  # Activation volume (mL/mol)
    rate_modifier = math.exp(-delta_v * (pressure - 1) * 0.001 / (8.314 * (temperature + 273.15)))

    # Yield estimate
    base_yield = 60 + (seed % 30)
    temp_factor = 1.0 if 20 <= temperature <= 80 else (0.8 if temperature < 20 else 0.7)
    est_yield = min(99, round(base_yield * temp_factor * min(rate_modifier, 1.3), 1))

    return {
        "Temperature_C": temperature,
        "Pressure_atm": pressure,
        "Thermal_Stability": temp_stability,
        "Decomposition_Temp_Est": f"~{decomp_temp}°C",
        "Rate_Modifier": round(rate_modifier, 3),
        "Estimated_Yield": f"{est_yield}%",
        "Activation_Volume": f"{delta_v} mL/mol",
        "Recommendation": "Optimal" if 20 <= temperature <= 60 and 0.8 <= pressure <= 2 else "Adjust conditions"
    }


# ═══════════════════════════════════════════════════════════════════
# 2. SOLVENT EFFECT SIMULATION
# ═══════════════════════════════════════════════════════════════════

SOLVENTS = {
    "Water":       {"polarity": 10.2, "bp": 100, "dielectric": 80.1},
    "DMSO":        {"polarity": 7.2,  "bp": 189, "dielectric": 46.7},
    "DMF":         {"polarity": 6.4,  "bp": 153, "dielectric": 36.7},
    "Ethanol":     {"polarity": 5.2,  "bp": 78,  "dielectric": 24.3},
    "Methanol":    {"polarity": 5.1,  "bp": 65,  "dielectric": 32.7},
    "Acetone":     {"polarity": 5.1,  "bp": 56,  "dielectric": 20.7},
    "THF":         {"polarity": 4.0,  "bp": 66,  "dielectric": 7.6},
    "DCM":         {"polarity": 3.1,  "bp": 40,  "dielectric": 8.9},
    "Chloroform":  {"polarity": 2.7,  "bp": 61,  "dielectric": 4.8},
    "Toluene":     {"polarity": 2.4,  "bp": 111, "dielectric": 2.4},
    "Hexane":      {"polarity": 0.1,  "bp": 69,  "dielectric": 1.9},
}


def solvent_effect_simulation(mol):
    """Compare reaction outcomes across different solvents."""
    logp = Descriptors.MolLogP(mol)
    tpsa = CalcTPSA(mol)
    seed = _mol_seed(mol)
    results = []

    for name, props in SOLVENTS.items():
        # Solubility compatibility
        polarity_match = abs(logp - props["polarity"] * 0.5)
        if polarity_match < 2:
            solubility = "High"
        elif polarity_match < 4:
            solubility = "Moderate"
        else:
            solubility = "Low"

        # Reaction rate effect
        rate_factor = 1.0 + (props["dielectric"] * 0.005) - (abs(logp) * 0.02)
        rate_factor = max(0.5, min(2.0, rate_factor))

        # Yield estimate
        base = 50 + (seed % 30)
        y = min(99, round(base * (1 if solubility == "High" else 0.7 if solubility == "Moderate" else 0.4) * min(rate_factor, 1.5), 1))

        results.append({
            "Solvent": name,
            "Polarity": props["polarity"],
            "Boiling_Point": f"{props['bp']}°C",
            "Dielectric": props["dielectric"],
            "Solubility": solubility,
            "Rate_Factor": round(rate_factor, 2),
            "Est_Yield": f"{y}%",
        })

    # Sort by yield
    results.sort(key=lambda x: float(x["Est_Yield"].replace("%", "")), reverse=True)
    return results


# ═══════════════════════════════════════════════════════════════════
# 3. CATALYST TESTING MODE
# ═══════════════════════════════════════════════════════════════════

CATALYSTS = [
    {"name": "Pd/C", "type": "Heterogeneous", "selectivity": "High", "cost": "High"},
    {"name": "Pd(PPh₃)₄", "type": "Homogeneous", "selectivity": "Very High", "cost": "Very High"},
    {"name": "RuCl₃", "type": "Homogeneous", "selectivity": "Medium", "cost": "Medium"},
    {"name": "FeCl₃", "type": "Lewis Acid", "selectivity": "Low", "cost": "Low"},
    {"name": "AlCl₃", "type": "Lewis Acid", "selectivity": "Medium", "cost": "Low"},
    {"name": "NaOH", "type": "Base", "selectivity": "Low", "cost": "Very Low"},
    {"name": "H₂SO₄", "type": "Acid", "selectivity": "Low", "cost": "Very Low"},
    {"name": "Grubbs Gen2", "type": "Metathesis", "selectivity": "Very High", "cost": "Very High"},
    {"name": "CuI", "type": "Click Chemistry", "selectivity": "High", "cost": "Medium"},
    {"name": "Ni(COD)₂", "type": "Cross-Coupling", "selectivity": "High", "cost": "High"},
]


def catalyst_testing(mol):
    """Simulate catalyst influence on reaction rate and selectivity."""
    seed = _mol_seed(mol)
    mw = Descriptors.MolWt(mol)
    results = []

    for i, cat in enumerate(CATALYSTS):
        base_rate = 1.0 + ((seed + i * 137) % 50) / 25
        # Selectivity boost
        sel_map = {"Very High": 1.5, "High": 1.3, "Medium": 1.1, "Low": 0.9}
        sel_factor = sel_map.get(cat["selectivity"], 1.0)

        # Yield
        y = min(99, round(40 + (seed % 20) + base_rate * sel_factor * 15, 1))

        # Side products
        side_products = max(0, round(100 - y - 5 * sel_factor, 1))

        results.append({
            "Catalyst": cat["name"],
            "Type": cat["type"],
            "Selectivity": cat["selectivity"],
            "Rate_Enhancement": f"{round(base_rate * sel_factor, 2)}x",
            "Est_Yield": f"{y}%",
            "Side_Products": f"{side_products}%",
            "Cost": cat["cost"],
        })

    results.sort(key=lambda x: float(x["Est_Yield"].replace("%", "")), reverse=True)
    return results


# ═══════════════════════════════════════════════════════════════════
# 4. pH VARIATION TESTING
# ═══════════════════════════════════════════════════════════════════

def ph_variation_testing(mol):
    """Test reaction outcomes across pH 1-14."""
    logp = Descriptors.MolLogP(mol)
    seed = _mol_seed(mol)
    has_acid = mol.HasSubstructMatch(Chem.MolFromSmarts("C(=O)O"))
    has_amine = mol.HasSubstructMatch(Chem.MolFromSmarts("[NX3;H1,H2;!$(N-C=O)]"))

    results = []
    for ph in range(1, 15):
        # Ionization state
        if has_acid and has_amine:
            if ph < 4:
                ionization = "Cationic"
            elif ph > 9:
                ionization = "Anionic"
            else:
                ionization = "Zwitterionic"
        elif has_acid:
            ionization = "Anionic" if ph > 5 else "Neutral"
        elif has_amine:
            ionization = "Cationic" if ph < 8 else "Neutral"
        else:
            ionization = "Neutral"

        # Solubility at this pH
        if ionization in ("Cationic", "Anionic", "Zwitterionic"):
            sol = "High"
        elif logp > 3:
            sol = "Low"
        else:
            sol = "Moderate"

        # Stability
        stability = 100
        if ph < 2 or ph > 12:
            stability -= 40  # extreme pH
        if has_acid and ph > 10:
            stability -= 20  # base hydrolysis of esters
        stability = max(0, stability + (seed % 15))

        # Absorption
        if ionization == "Neutral":
            absorption = "Best"
        elif ionization == "Zwitterionic":
            absorption = "Good"
        else:
            absorption = "Limited"

        results.append({
            "pH": ph,
            "Ionization": ionization,
            "Solubility": sol,
            "Stability_%": min(100, stability),
            "Absorption": absorption,
        })

    return results


# ═══════════════════════════════════════════════════════════════════
# 5. REACTION RATE SIMULATION (Kinetics)
# ═══════════════════════════════════════════════════════════════════

def reaction_rate_simulation(mol, temperatures=None):
    """Approximate kinetics using Arrhenius model."""
    if temperatures is None:
        temperatures = [10, 20, 25, 30, 40, 50, 60, 80, 100]

    seed = _mol_seed(mol)
    mw = Descriptors.MolWt(mol)

    # Activation energy estimate (kJ/mol)
    Ea = 40 + (seed % 60)  # 40-100 kJ/mol range
    A = 1e10 + (seed % 1000) * 1e7  # Pre-exponential factor

    R = 8.314  # J/(mol·K)
    results = []

    for T_c in temperatures:
        T_k = T_c + 273.15
        k = A * math.exp(-Ea * 1000 / (R * T_k))  # Rate constant
        t_half = round(0.693 / k, 4) if k > 0 else float('inf')

        results.append({
            "Temperature_C": T_c,
            "Temperature_K": round(T_k, 1),
            "Rate_Constant_k": f"{k:.4e}",
            "Half_Life_s": t_half if t_half < 1e6 else ">1e6",
        })

    return {
        "Activation_Energy_kJ": Ea,
        "Pre_Exponential": f"{A:.2e}",
        "Kinetic_Order": "First Order (assumed)",
        "Data": results,
    }


# ═══════════════════════════════════════════════════════════════════
# 6. EQUILIBRIUM TESTING MODE
# ═══════════════════════════════════════════════════════════════════

def equilibrium_testing(mol):
    """Visualize forward vs reverse reaction equilibrium."""
    seed = _mol_seed(mol)
    logp = Descriptors.MolLogP(mol)
    mw = Descriptors.MolWt(mol)

    # Equilibrium constant
    delta_G = -20 + (seed % 40)  # kJ/mol
    R = 8.314
    T = 298.15
    K = math.exp(-delta_G * 1000 / (R * T))

    # Concentrations at equilibrium
    if K > 100:
        direction = "Products Strongly Favored"
        product_pct = round(95 + (seed % 5), 1)
    elif K > 1:
        direction = "Products Favored"
        product_pct = round(60 + (seed % 30), 1)
    elif K > 0.01:
        direction = "Reactants Favored"
        product_pct = round(10 + (seed % 30), 1)
    else:
        direction = "Reactants Strongly Favored"
        product_pct = round(1 + (seed % 8), 1)

    return {
        "Delta_G_kJ": delta_G,
        "Equilibrium_K": round(K, 4) if K < 1e6 else f"{K:.2e}",
        "Direction": direction,
        "Product_Pct": product_pct,
        "Reactant_Pct": round(100 - product_pct, 1),
        "Temperature_K": T,
        "Le_Chatelier_Shift": "Increase T → shifts to endothermic side" if delta_G > 0 else "Increase T → shifts to reactant side",
    }


# ═══════════════════════════════════════════════════════════════════
# 7. MULTI-REAGENT COMPATIBILITY TESTER
# ═══════════════════════════════════════════════════════════════════

def multi_reagent_compatibility(mol):
    """Test compatibility with common reagents."""
    reagents = [
        ("NaBH₄", "Reducing Agent", ["[CX3H1](=O)", "[CX3](=O)[CX3]"]),
        ("LiAlH₄", "Strong Reducer", ["[CX3](=O)", "[CX3](=O)[OX2]"]),
        ("mCPBA", "Oxidizing Agent", ["C=C", "[SX2]"]),
        ("H₂/Pd", "Hydrogenation", ["C=C", "C#C", "[N+](=O)[O-]"]),
        ("NaOH", "Base Hydrolysis", ["C(=O)OC", "C(=O)Cl"]),
        ("HCl", "Acid Conditions", ["[NX3]", "C1OC1"]),
        ("SOCl₂", "Chlorination", ["[OH]", "C(=O)O"]),
        ("Grignard (RMgBr)", "Nucleophilic", ["[CX3H1](=O)", "[CX3](=O)[CX3]", "C1OC1"]),
    ]

    results = []
    for rname, rtype, patterns in reagents:
        reactive_groups = []
        for p in patterns:
            pat = Chem.MolFromSmarts(p)
            if pat and mol.HasSubstructMatch(pat):
                reactive_groups.append(p)

        compatible = len(reactive_groups) > 0
        results.append({
            "Reagent": rname,
            "Type": rtype,
            "Reactive_Groups_Found": len(reactive_groups),
            "Compatible": "✅ Yes" if compatible else "❌ No reactive groups",
            "Expected_Result": f"Reaction at {len(reactive_groups)} site(s)" if compatible else "No reaction expected",
        })

    return results


# ═══════════════════════════════════════════════════════════════════
# 8. THERMODYNAMIC STABILITY TESTING
# ═══════════════════════════════════════════════════════════════════

def thermodynamic_stability(mol):
    """Assess thermodynamic stability of the molecule."""
    seed = _mol_seed(mol)
    mw = Descriptors.MolWt(mol)
    rings = CalcNumRings(mol)
    arom = CalcNumAromaticRings(mol)

    # Heat of formation estimate
    hof = -100 + (seed % 200) + mw * 0.1 - arom * 30  # kJ/mol (rough)

    # Bond dissociation energies
    num_bonds = mol.GetNumBonds()
    avg_bde = round(350 + (seed % 100), 0)  # kJ/mol average

    # Entropy estimate
    entropy = round(200 + mw * 0.5 + CalcNumRotatableBonds(mol) * 10, 1)

    # Gibbs free energy
    T = 298.15
    gibbs = round(hof - T * entropy / 1000, 1)

    # Overall stability
    if gibbs < -200:
        stability = "Very Stable"
    elif gibbs < -50:
        stability = "Stable"
    elif gibbs < 50:
        stability = "Moderately Stable"
    else:
        stability = "Unstable"

    return {
        "Heat_of_Formation_kJ": round(hof, 1),
        "Avg_BDE_kJ": avg_bde,
        "Entropy_JmolK": entropy,
        "Gibbs_Free_Energy_kJ": gibbs,
        "Stability_Class": stability,
        "Aromatic_Stabilization": f"{arom} aromatic rings (thermodynamically favorable)",
        "Ring_Strain": "Present" if any(len(r) <= 4 for r in mol.GetRingInfo().AtomRings()) else "Absent",
    }


# ═══════════════════════════════════════════════════════════════════
# 9. REACTION SENSITIVITY ANALYSIS
# ═══════════════════════════════════════════════════════════════════

def sensitivity_analysis(mol, parameter="temperature"):
    """Small change in conditions → output change analysis."""
    seed = _mol_seed(mol)
    base_yield = 60 + (seed % 25)

    if parameter == "temperature":
        variations = list(range(-20, 25, 5))
        label = "ΔT (°C)"
    elif parameter == "concentration":
        variations = [-50, -25, -10, 0, 10, 25, 50]
        label = "ΔConc (%)"
    elif parameter == "pressure":
        variations = [-2, -1, -0.5, 0, 0.5, 1, 2]
        label = "ΔP (atm)"
    else:
        variations = list(range(-5, 6))
        label = "ΔpH"

    results = []
    for delta in variations:
        sensitivity = 0.3 + (seed % 10) / 10  # sensitivity coefficient
        yield_change = round(delta * sensitivity, 2)
        new_yield = max(0, min(100, round(base_yield + yield_change, 1)))

        results.append({
            label: delta,
            "Base_Yield": base_yield,
            "Yield_Change": yield_change,
            "New_Yield": new_yield,
        })

    return {
        "Parameter": parameter,
        "Sensitivity_Coefficient": round(sensitivity, 3),
        "Classification": "Sensitive" if sensitivity > 0.8 else ("Moderate" if sensitivity > 0.4 else "Robust"),
        "Data": results,
    }


# ═══════════════════════════════════════════════════════════════════
# 10. SIDE REACTION EXPLORATION
# ═══════════════════════════════════════════════════════════════════

def side_reaction_exploration(mol):
    """Identify potential side reactions."""
    side_reactions = []

    checks = [
        ("C=C", "Polymerization", "Alkene may undergo radical polymerization at high temp"),
        ("c1ccccc1O", "Oxidative Coupling", "Phenol can form dimers via radical oxidation"),
        ("[NX3;H2]", "Diazonium Formation", "Primary amine → diazo salt with HNO₂"),
        ("C(=O)O", "Decarboxylation", "Carboxylic acid may lose CO₂ at high temperature"),
        ("[CH3]O", "Demethylation", "Methyl ether can be cleaved by strong acids (BBr₃)"),
        ("C#C", "Cycloaddition", "Alkyne susceptible to [2+2] cycloaddition under UV"),
        ("C(=O)Cl", "Hydrolysis", "Acyl chloride rapidly hydrolyzes in water"),
        ("[OH]", "Elimination", "Alcohol may undergo E1/E2 elimination"),
        ("[NX3][CX3]=O", "Amide Hydrolysis", "Amide bond may cleave under harsh acidic/basic conditions"),
    ]

    for smarts, rxn_name, description in checks:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            count = len(mol.GetSubstructMatches(pat))
            side_reactions.append({
                "Reaction": rxn_name,
                "Sites": count,
                "Description": description,
                "Risk": "High" if count > 2 else "Moderate",
            })

    if not side_reactions:
        side_reactions.append({
            "Reaction": "None Detected",
            "Sites": 0,
            "Description": "No common side-reaction motifs found",
            "Risk": "Low",
        })

    return side_reactions


# ═══════════════════════════════════════════════════════════════════
# 11. CHEMICAL DEGRADATION TESTING
# ═══════════════════════════════════════════════════════════════════

def degradation_testing(mol):
    """Predict degradation pathways."""
    seed = _mol_seed(mol)
    pathways = []

    degradation_checks = [
        ("C(=O)OC", "Ester Hydrolysis", "Hydrolytic", "Weeks-Months", "Neutral pH, water"),
        ("[NX3][CX3](=O)", "Amide Hydrolysis", "Hydrolytic", "Months-Years", "Strong acid/base"),
        ("C=C", "Oxidative Degradation", "Oxidative", "Days-Weeks", "Air, light exposure"),
        ("[SX2]", "Sulfide Oxidation", "Oxidative", "Days", "Air exposure, H₂O₂"),
        ("c1ccccc1O", "Phenol Oxidation", "Oxidative", "Hours-Days", "Air + metal ions"),
        ("[CX3H1](=O)", "Aldehyde Oxidation", "Oxidative", "Hours", "Air exposure → carboxylic acid"),
        ("C#C", "Alkyne Polymerization", "Thermal", "Months", "Elevated temperature"),
    ]

    for smarts, name, mechanism, t_half, conditions in degradation_checks:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            pathways.append({
                "Pathway": name,
                "Mechanism": mechanism,
                "Estimated_Half_Life": t_half,
                "Triggering_Conditions": conditions,
            })

    if not pathways:
        pathways.append({
            "Pathway": "No degradation hotspots",
            "Mechanism": "N/A",
            "Estimated_Half_Life": ">1 year",
            "Triggering_Conditions": "Chemically robust structure",
        })

    return pathways


# ═══════════════════════════════════════════════════════════════════
# 12. ENVIRONMENTAL CONDITION TESTING
# ═══════════════════════════════════════════════════════════════════

def environmental_condition_testing(mol):
    """Test stability under humidity, temperature, light."""
    seed = _mol_seed(mol)
    logp = Descriptors.MolLogP(mol)
    mw = Descriptors.MolWt(mol)

    # Hygroscopic tendency
    tpsa = CalcTPSA(mol)
    hbd = CalcNumHBD(mol)
    hygroscopic = "High" if (tpsa > 120 and hbd > 3) else ("Moderate" if tpsa > 80 else "Low")

    # Photostability
    arom = CalcNumAromaticRings(mol)
    has_conjugation = mol.HasSubstructMatch(Chem.MolFromSmarts("C=CC=C"))
    photostable = "Low" if (arom > 2 or has_conjugation) else ("Moderate" if arom > 0 else "High")

    # Thermal stability
    decomp = 150 + seed % 200
    thermal = "Stable (<150°C)" if decomp > 200 else ("Moderate" if decomp > 100 else "Labile")

    # Oxidative sensitivity
    has_thiol = mol.HasSubstructMatch(Chem.MolFromSmarts("[SH]"))
    has_phenol = mol.HasSubstructMatch(Chem.MolFromSmarts("c1ccccc1O"))
    oxidative = "High" if (has_thiol or has_phenol) else "Low"

    return {
        "Hygroscopic_Tendency": hygroscopic,
        "Photostability": photostable,
        "Thermal_Stability": thermal,
        "Decomposition_Temp_Est": f"~{decomp}°C",
        "Oxidative_Sensitivity": oxidative,
        "Recommended_Storage": "Cool, dry, dark place" if any(x != "Low" for x in [hygroscopic, oxidative]) else "Room temperature, no special requirements",
        "Shelf_Life_Est": "6-12 months" if photostable == "Low" or oxidative == "High" else "2-5 years",
    }


# ═══════════════════════════════════════════════════════════════════
# 13. CONCENTRATION VARIATION SIMULATION
# ═══════════════════════════════════════════════════════════════════

def concentration_variation(mol, concentrations=None):
    """Simulate reaction outcomes at different concentrations."""
    if concentrations is None:
        concentrations = [0.01, 0.05, 0.1, 0.25, 0.5, 1.0, 2.0, 5.0]  # mol/L

    seed = _mol_seed(mol)
    results = []

    for conc in concentrations:
        # Rate is proportional to concentration (first-order)
        rate = round(conc * (1.0 + (seed % 50) / 100), 4)
        # Yield optimization curve (bell-shaped)
        optimal_conc = 0.5 + (seed % 10) / 10
        yield_val = max(20, min(99, round(90 - 30 * abs(math.log(conc / optimal_conc)), 1)))

        # Selectivity (decreases at very high concentration)
        selectivity = round(95 - conc * 5, 1) if conc < 5 else 70

        results.append({
            "Concentration_M": conc,
            "Reaction_Rate": rate,
            "Est_Yield_%": yield_val,
            "Selectivity_%": selectivity,
        })

    return {
        "Optimal_Concentration_M": round(optimal_conc, 2),
        "Data": results,
    }


# ═══════════════════════════════════════════════════════════════════
# 14. REACTION REPEATABILITY TESTING
# ═══════════════════════════════════════════════════════════════════

def repeatability_testing(mol, n_runs=10):
    """Simulate n runs and calculate reproducibility stats."""
    seed = _mol_seed(mol)
    np.random.seed(seed % 10000)

    base_yield = 60 + (seed % 25)
    std_dev = 2 + (seed % 8)  # Higher = less reproducible

    yields = np.random.normal(base_yield, std_dev, n_runs)
    yields = np.clip(yields, 0, 100).round(1)

    cv = round(np.std(yields) / np.mean(yields) * 100, 2) if np.mean(yields) > 0 else 0

    return {
        "Runs": n_runs,
        "Mean_Yield": round(float(np.mean(yields)), 1),
        "Std_Dev": round(float(np.std(yields)), 2),
        "Min_Yield": float(np.min(yields)),
        "Max_Yield": float(np.max(yields)),
        "CV_Pct": cv,
        "Reproducibility": "Excellent" if cv < 3 else ("Good" if cv < 5 else ("Fair" if cv < 10 else "Poor")),
        "Individual_Runs": yields.tolist(),
    }


# ═══════════════════════════════════════════════════════════════════
# 15. ERROR SENSITIVITY TESTING
# ═══════════════════════════════════════════════════════════════════

def error_sensitivity_testing(mol):
    """Assess how sensitive the output is to small input changes."""
    seed = _mol_seed(mol)
    logp = Descriptors.MolLogP(mol)
    mw = Descriptors.MolWt(mol)
    tpsa = CalcTPSA(mol)

    parameters = {
        "Temperature (±5°C)": round(0.5 + (seed % 20) / 10, 2),
        "Concentration (±10%)": round(0.3 + (seed % 15) / 10, 2),
        "Pressure (±0.5 atm)": round(0.1 + (seed % 10) / 10, 2),
        "pH (±0.5)": round(0.4 + (seed % 12) / 10, 2),
        "Reaction Time (±10%)": round(0.2 + (seed % 8) / 10, 2),
        "Catalyst Loading (±5%)": round(0.6 + (seed % 18) / 10, 2),
    }

    results = []
    for param, sensitivity in parameters.items():
        yield_impact = round(sensitivity * 5, 1)  # % yield change
        classification = "Critical" if sensitivity > 1.5 else ("Sensitive" if sensitivity > 0.8 else "Robust")
        results.append({
            "Parameter": param,
            "Sensitivity_Index": sensitivity,
            "Yield_Impact_%": yield_impact,
            "Classification": classification,
        })

    # Sort by sensitivity
    results.sort(key=lambda x: x["Sensitivity_Index"], reverse=True)

    return {
        "Overall_Robustness": "Robust" if all(r["Classification"] == "Robust" for r in results)
            else ("Moderate" if sum(1 for r in results if r["Classification"] == "Critical") <= 1
            else "Fragile"),
        "Parameters": results,
    }
