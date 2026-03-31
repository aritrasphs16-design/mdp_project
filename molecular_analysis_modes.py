"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   CHEMOFILTER — MOLECULAR ANALYSIS MODES (16-25)                            ║
║   Bond Strength · Geometry · Steric · Electron Density · H-Bonds · VdW     ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

from rdkit import Chem
from rdkit.Chem import Descriptors, AllChem, Crippen
from rdkit.Chem.rdMolDescriptors import (
    CalcNumHBD, CalcNumHBA, CalcNumRotatableBonds,
    CalcNumRings, CalcNumAromaticRings, CalcTPSA,
    CalcLabuteASA, CalcFractionCSP3,
    CalcKappa1, CalcKappa2, CalcKappa3,
    CalcNumHeteroatoms
)
import math
import hashlib


def _seed(mol):
    return int(hashlib.md5(Chem.MolToSmiles(mol).encode()).hexdigest(), 16) % 10**8


# ═══════════════════════════════════════════════════════════════════
# 16. BOND STRENGTH ESTIMATION
# ═══════════════════════════════════════════════════════════════════

# Average bond dissociation energies (kJ/mol)
BDE_TABLE = {
    ("C", "C", 1.0): 346, ("C", "C", 1.5): 518, ("C", "C", 2.0): 614,
    ("C", "C", 3.0): 839,
    ("C", "H", 1.0): 411, ("C", "O", 1.0): 358, ("C", "O", 1.5): 460,
    ("C", "O", 2.0): 799, ("C", "N", 1.0): 305, ("C", "N", 1.5): 410,
    ("C", "N", 2.0): 615, ("C", "N", 3.0): 891,
    ("C", "F", 1.0): 485, ("C", "Cl", 1.0): 339, ("C", "Br", 1.0): 285,
    ("C", "I", 1.0): 213, ("C", "S", 1.0): 272, ("C", "S", 2.0): 577,
    ("O", "H", 1.0): 459, ("N", "H", 1.0): 386, ("N", "N", 1.0): 167,
    ("N", "N", 2.0): 418, ("N", "N", 3.0): 942, ("O", "O", 1.0): 142,
    ("O", "O", 2.0): 498, ("S", "H", 1.0): 363, ("S", "S", 1.0): 226,
    ("N", "O", 1.0): 201, ("N", "O", 2.0): 607,
    ("P", "O", 1.0): 335, ("P", "O", 2.0): 544,
}


def bond_strength_estimation(mol):
    """Estimate bond dissociation energies for all bonds in the molecule."""
    bond_data = []
    weakest = None
    weakest_bde = float('inf')

    for bond in mol.GetBonds():
        a1 = bond.GetBeginAtom().GetSymbol()
        a2 = bond.GetEndAtom().GetSymbol()
        bt = bond.GetBondTypeAsDouble()

        # Lookup in table (try both orderings)
        bde = BDE_TABLE.get((a1, a2, bt)) or BDE_TABLE.get((a2, a1, bt))
        if bde is None:
            bde = 350  # Default estimate

        # Classify
        if bde < 250:
            strength = "Weak"
        elif bde < 400:
            strength = "Moderate"
        else:
            strength = "Strong"

        bond_info = {
            "Bond": f"{a1}-{a2}",
            "Type": str(bond.GetBondType().name),
            "BDE_kJ_mol": bde,
            "Strength": strength,
            "Idx": bond.GetIdx(),
        }
        bond_data.append(bond_info)

        if bde < weakest_bde:
            weakest_bde = bde
            weakest = bond_info

    # Summary
    avg_bde = round(sum(b["BDE_kJ_mol"] for b in bond_data) / len(bond_data), 1) if bond_data else 0

    return {
        "Total_Bonds": len(bond_data),
        "Average_BDE": avg_bde,
        "Weakest_Bond": weakest,
        "Strongest_Bonds": sorted(bond_data, key=lambda x: x["BDE_kJ_mol"], reverse=True)[:3],
        "All_Bonds": bond_data[:20],  # Limit for display
    }


# ═══════════════════════════════════════════════════════════════════
# 17. MOLECULAR GEOMETRY ANALYZER
# ═══════════════════════════════════════════════════════════════════

def molecular_geometry_analyzer(mol):
    """Analyze molecular geometry — bond lengths, angles, shape."""
    heavy = mol.GetNumHeavyAtoms()
    rings = CalcNumRings(mol)
    arom = CalcNumAromaticRings(mol)
    rot = CalcNumRotatableBonds(mol)
    fsp3 = CalcFractionCSP3(mol)

    # Shape analysis
    if fsp3 > 0.6:
        shape = "Globular (3D)"
    elif fsp3 > 0.3:
        shape = "Intermediate"
    else:
        shape = "Planar (2D)"

    # Kappa shape indices
    try:
        k1 = round(CalcKappa1(mol), 3)
        k2 = round(CalcKappa2(mol), 3)
        k3 = round(CalcKappa3(mol), 3)
    except:
        k1 = k2 = k3 = 0

    # Hybridization analysis
    sp3 = sum(1 for a in mol.GetAtoms() if a.GetHybridization().name == "SP3")
    sp2 = sum(1 for a in mol.GetAtoms() if a.GetHybridization().name == "SP2")
    sp = sum(1 for a in mol.GetAtoms() if a.GetHybridization().name == "SP")

    # Planarity score
    planarity = round((sp2 + sp) / heavy * 100, 1) if heavy > 0 else 0

    # Surface area
    asa = round(CalcLabuteASA(mol), 1)

    return {
        "Heavy_Atoms": heavy,
        "Molecular_Shape": shape,
        "Planarity_Score": f"{planarity}%",
        "Fsp3": round(fsp3, 3),
        "Kappa1": k1,
        "Kappa2": k2,
        "Kappa3": k3,
        "sp3_Centers": sp3,
        "sp2_Centers": sp2,
        "sp_Centers": sp,
        "Rings": rings,
        "Aromatic_Rings": arom,
        "Rotatable_Bonds": rot,
        "Approx_Surface_Area": f"{asa} Å²",
    }


# ═══════════════════════════════════════════════════════════════════
# 18. STERIC HINDRANCE DETECTION
# ═══════════════════════════════════════════════════════════════════

def steric_hindrance_detection(mol):
    """Detect sterically hindered sites."""
    hindered_sites = []

    # Quaternary carbons
    quat_c = Chem.MolFromSmarts("[CX4]([#6])([#6])([#6])[#6]")
    if quat_c:
        matches = mol.GetSubstructMatches(quat_c)
        if matches:
            hindered_sites.append({
                "Type": "Quaternary Carbon",
                "Count": len(matches),
                "Impact": "High steric bulk, limits access",
            })

    # gem-Dimethyl
    gem_di = Chem.MolFromSmarts("[CX4]([CH3])([CH3])")
    if gem_di:
        matches = mol.GetSubstructMatches(gem_di)
        if matches:
            hindered_sites.append({
                "Type": "gem-Dimethyl",
                "Count": len(matches),
                "Impact": "Thorpe-Ingold effect, cyclization favored",
            })

    # tert-Butyl
    tbu = Chem.MolFromSmarts("C(C)(C)(C)")
    if tbu:
        matches = mol.GetSubstructMatches(tbu)
        if matches:
            hindered_sites.append({
                "Type": "tert-Butyl Group",
                "Count": len(matches),
                "Impact": "Major steric shield, metabolic block",
            })

    # Ortho-substituted aromatics
    ortho = Chem.MolFromSmarts("c1c([!H])c([!H])ccc1")
    if ortho:
        matches = mol.GetSubstructMatches(ortho)
        if matches:
            hindered_sites.append({
                "Type": "Ortho-Disubstituted Arene",
                "Count": len(matches),
                "Impact": "Restricted rotation, atropisomerism possible",
            })

    # Neopentyl
    neo = Chem.MolFromSmarts("[CH2]C(C)(C)C")
    if neo:
        matches = mol.GetSubstructMatches(neo)
        if matches:
            hindered_sites.append({
                "Type": "Neopentyl",
                "Count": len(matches),
                "Impact": "SN2 reactions disfavored",
            })

    # Overall assessment
    total_hindrance = sum(s["Count"] for s in hindered_sites)
    if total_hindrance == 0:
        overall = "Low Steric Hindrance"
    elif total_hindrance <= 2:
        overall = "Moderate Steric Hindrance"
    else:
        overall = "High Steric Hindrance"

    return {
        "Overall": overall,
        "Total_Hindered_Sites": total_hindrance,
        "Details": hindered_sites if hindered_sites else [{"Type": "None detected", "Count": 0, "Impact": "Open topology"}],
    }


# ═══════════════════════════════════════════════════════════════════
# 19. ELECTRON DENSITY DISTRIBUTION ANALYSIS
# ═══════════════════════════════════════════════════════════════════

def electron_density_analysis(mol):
    """Analyze electron density distribution using Gasteiger charges."""
    try:
        AllChem.ComputeGasteigerCharges(mol)
    except:
        return {"Error": "Could not compute Gasteiger charges"}

    charges = []
    most_positive = None
    most_negative = None
    max_q = -999
    min_q = 999

    for atom in mol.GetAtoms():
        q = float(atom.GetProp("_GasteigerCharge"))
        if math.isnan(q) or math.isinf(q):
            q = 0.0
        charges.append({
            "Atom": atom.GetSymbol(),
            "Idx": atom.GetIdx(),
            "Charge": round(q, 4),
            "Type": "Electrophilic" if q > 0.1 else ("Nucleophilic" if q < -0.1 else "Neutral"),
        })
        if q > max_q:
            max_q = q
            most_positive = charges[-1]
        if q < min_q:
            min_q = q
            most_negative = charges[-1]

    # Dipole moment estimate
    dipole = round(sum(abs(c["Charge"]) for c in charges) / len(charges) * 10, 2) if charges else 0

    return {
        "Most_Electrophilic": most_positive,
        "Most_Nucleophilic": most_negative,
        "Charge_Range": round(max_q - min_q, 4),
        "Dipole_Estimate": dipole,
        "Polarized_Atoms": len([c for c in charges if abs(c["Charge"]) > 0.1]),
        "Atom_Charges": charges[:30],  # Limit for display
    }


# ═══════════════════════════════════════════════════════════════════
# 20. CHARGE DISTRIBUTION VISUALIZATION DATA
# ═══════════════════════════════════════════════════════════════════

def charge_distribution_data(mol):
    """Prepare charge distribution data for visualization."""
    try:
        AllChem.ComputeGasteigerCharges(mol)
    except:
        return {"atoms": [], "charge_histogram": []}

    atoms = []
    charge_bins = {f"{i/10:.1f}": 0 for i in range(-10, 11)}

    for atom in mol.GetAtoms():
        q = float(atom.GetProp("_GasteigerCharge"))
        if math.isnan(q) or math.isinf(q):
            q = 0.0
        atoms.append({
            "symbol": atom.GetSymbol(),
            "idx": atom.GetIdx(),
            "charge": round(q, 3),
        })
        # Bin the charge
        bin_key = f"{round(q, 1):.1f}"
        if bin_key in charge_bins:
            charge_bins[bin_key] += 1

    return {
        "atoms": atoms,
        "charge_histogram": [{"range": k, "count": v} for k, v in charge_bins.items() if v > 0],
    }


# ═══════════════════════════════════════════════════════════════════
# 21. FUNCTIONAL GROUP REACTIVITY PREDICTION
# ═══════════════════════════════════════════════════════════════════

def functional_group_reactivity(mol):
    """Predict reactivity of each functional group found."""
    fg_reactivity = [
        ("[OX2H1]", "Hydroxyl (-OH)", "Moderate", "Substitution, oxidation, H-bonding"),
        ("[CX3](=O)[OX2H1]", "Carboxylic Acid", "High", "Esterification, salt formation, decarboxylation"),
        ("[NX3;H2;!$(N-C=O)]", "Primary Amine", "High", "Alkylation, acylation, diazotization"),
        ("[NX3;H1;!$(N-C=O)]", "Secondary Amine", "Moderate", "Alkylation, Mannich reaction"),
        ("[NX3][CX3](=[OX1])", "Amide", "Low", "Hydrolysis (harsh), N-alkylation"),
        ("[CX3](=O)[OX2H0]", "Ester", "Moderate", "Hydrolysis, transesterification"),
        ("[CX3H1](=O)", "Aldehyde", "Very High", "Nucleophilic addition, oxidation, Wittig"),
        ("[CX3](=O)([#6])[#6]", "Ketone", "Moderate", "Nucleophilic addition, enolization"),
        ("[F,Cl,Br,I][CX4]", "Alkyl Halide", "High", "SN1/SN2, elimination"),
        ("C=C", "Alkene", "Moderate", "Electrophilic addition, polymerization"),
        ("C#C", "Alkyne", "Moderate", "Hydrogenation, click chemistry"),
        ("[CX2]#N", "Nitrile", "Low-Moderate", "Hydrolysis to amide/acid"),
        ("[SX2H]", "Thiol", "High", "Disulfide formation, alkylation"),
        ("[N+](=O)[O-]", "Nitro", "Low", "Reduction to amine"),
        ("c1ccccc1", "Aromatic Ring", "Low", "Electrophilic substitution"),
    ]

    found = []
    for smarts, name, reactivity, reactions in fg_reactivity:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            count = len(mol.GetSubstructMatches(pat))
            found.append({
                "Group": name,
                "Count": count,
                "Reactivity": reactivity,
                "Typical_Reactions": reactions,
            })

    if not found:
        found.append({"Group": "None detected", "Count": 0, "Reactivity": "N/A", "Typical_Reactions": "N/A"})

    return found


# ═══════════════════════════════════════════════════════════════════
# 22. HYDROGEN BOND DETECTION
# ═══════════════════════════════════════════════════════════════════

def hydrogen_bond_detection(mol):
    """Detailed hydrogen bond donor/acceptor analysis."""
    hbd_total = CalcNumHBD(mol)
    hba_total = CalcNumHBA(mol)

    # Detailed donor analysis
    nh_count = len(mol.GetSubstructMatches(Chem.MolFromSmarts("[NH,NH2,NH3]")))
    oh_count = len(mol.GetSubstructMatches(Chem.MolFromSmarts("[OH]")))

    # Detailed acceptor analysis
    n_acceptors = len(mol.GetSubstructMatches(Chem.MolFromSmarts("[nX2,NX2,NX1]")))
    o_acceptors = len(mol.GetSubstructMatches(Chem.MolFromSmarts("[OX2,OX1]")))

    # Intramolecular H-bond potential
    # Crude heuristic: if donor and acceptor are 2-4 bonds apart
    intramolecular = "Possible" if (hbd_total > 0 and hba_total > 0 and CalcNumRings(mol) > 0) else "Unlikely"

    # H-bond balance
    ratio = hbd_total / hba_total if hba_total > 0 else float('inf')
    if 0.3 <= ratio <= 3.0:
        balance = "Balanced"
    elif ratio > 3.0:
        balance = "Donor-Heavy"
    else:
        balance = "Acceptor-Heavy"

    return {
        "Total_HBD": hbd_total,
        "Total_HBA": hba_total,
        "NH_Donors": nh_count,
        "OH_Donors": oh_count,
        "N_Acceptors": n_acceptors,
        "O_Acceptors": o_acceptors,
        "HBD_HBA_Ratio": round(ratio, 2) if ratio < 100 else ">>1",
        "Balance": balance,
        "Intramolecular_HBond": intramolecular,
        "Binding_Potential": "High" if hbd_total + hba_total >= 4 else ("Moderate" if hbd_total + hba_total >= 2 else "Low"),
    }


# ═══════════════════════════════════════════════════════════════════
# 23. VAN DER WAALS INTERACTION ANALYSIS
# ═══════════════════════════════════════════════════════════════════

VDW_RADII = {
    "H": 1.20, "C": 1.70, "N": 1.55, "O": 1.52,
    "F": 1.47, "Cl": 1.75, "Br": 1.85, "I": 1.98,
    "S": 1.80, "P": 1.80, "B": 1.92, "Si": 2.10,
}


def vdw_interaction_analysis(mol):
    """Analyze Van der Waals surface and interaction potential."""
    heavy = mol.GetNumHeavyAtoms()

    # Calculate total VdW volume (sum of atomic volumes)
    total_vdw_volume = 0
    atom_details = {}
    for atom in mol.GetAtoms():
        sym = atom.GetSymbol()
        r = VDW_RADII.get(sym, 1.70)
        vol = (4/3) * math.pi * r**3
        total_vdw_volume += vol
        atom_details[sym] = atom_details.get(sym, 0) + 1

    # Surface area estimate
    asa = CalcLabuteASA(mol)

    # Dispersion interactions (proportional to polarizability)
    mr = Crippen.MolMR(mol)
    dispersion_strength = "High" if mr > 100 else ("Moderate" if mr > 50 else "Low")

    # Hydrophobic contact area
    logp = Descriptors.MolLogP(mol)
    hydrophobic_sa = round(asa * (1 - CalcTPSA(mol) / asa) if asa > 0 else 0, 1)

    return {
        "Total_VdW_Volume": round(total_vdw_volume, 1),
        "Labute_ASA": round(asa, 1),
        "Molar_Refractivity": round(mr, 1),
        "Polarizability_Est": round(mr / 2.5, 2),
        "Dispersion_Strength": dispersion_strength,
        "Hydrophobic_SA": hydrophobic_sa,
        "Atom_Composition": atom_details,
        "Contact_Potential": "Favorable" if hydrophobic_sa > 100 else "Moderate",
    }


# ═══════════════════════════════════════════════════════════════════
# 24. MOLECULAR FLEXIBILITY ANALYSIS
# ═══════════════════════════════════════════════════════════════════

def molecular_flexibility_analysis(mol):
    """Comprehensive flexibility assessment."""
    rot = CalcNumRotatableBonds(mol)
    heavy = mol.GetNumHeavyAtoms()
    rings = CalcNumRings(mol)
    arom = CalcNumAromaticRings(mol)
    fsp3 = CalcFractionCSP3(mol)

    # Flexibility index
    flex_idx = round(rot / heavy * 100, 1) if heavy > 0 else 0

    # Rigidity from rings
    ring_fraction = round(sum(len(r) for r in mol.GetRingInfo().AtomRings()) / heavy * 100, 1) if heavy > 0 else 0

    # Conformational entropy estimate (more rotatable bonds = more entropy)
    conf_entropy = round(rot * 2.5, 1)  # kJ/mol·K approximate

    # Number of conformers estimate
    n_conformers = min(10000, 3 ** rot) if rot > 0 else 1

    # Classification
    if flex_idx > 30:
        classification = "Highly Flexible"
    elif flex_idx > 15:
        classification = "Moderately Flexible"
    elif flex_idx > 5:
        classification = "Semi-Rigid"
    else:
        classification = "Rigid"

    return {
        "Rotatable_Bonds": rot,
        "Flexibility_Index": flex_idx,
        "Ring_Fraction_%": ring_fraction,
        "Fsp3": round(fsp3, 3),
        "Conformational_Entropy": f"{conf_entropy} J/mol·K",
        "Est_Conformers": n_conformers,
        "Classification": classification,
        "Binding_Penalty": "High" if flex_idx > 30 else ("Moderate" if flex_idx > 15 else "Low"),
        "Recommendation": "Cyclization suggested" if flex_idx > 30 else "Acceptable flexibility",
    }


# ═══════════════════════════════════════════════════════════════════
# 25. CONFORMER COMPARISON MODE
# ═══════════════════════════════════════════════════════════════════

def conformer_comparison(mol, n_conformers=5):
    """Generate and compare multiple conformers."""
    try:
        m = Chem.AddHs(mol)
        cids = AllChem.EmbedMultipleConfs(m, numConfs=n_conformers, randomSeed=42, maxAttempts=100)

        if len(cids) == 0:
            return {"Error": "Could not generate conformers", "Conformers": []}

        conformers = []
        energies = []

        for cid in cids:
            try:
                props = AllChem.MMFFGetMoleculeProperties(m)
                ff = AllChem.MMFFGetMoleculeForceField(m, props, confId=cid)
                if ff:
                    ff.Minimize(maxIts=200)
                    energy = round(ff.CalcEnergy(), 2)
                else:
                    energy = 0.0
            except:
                energy = 0.0

            energies.append(energy)
            conformers.append({
                "Conformer_ID": cid,
                "Energy_kcal": energy,
            })

        # Relative energies
        min_e = min(energies) if energies else 0
        for c in conformers:
            c["Relative_Energy"] = round(c["Energy_kcal"] - min_e, 2)
            c["Population_Est"] = "Major" if c["Relative_Energy"] < 1 else (
                "Minor" if c["Relative_Energy"] < 3 else "Negligible"
            )

        conformers.sort(key=lambda x: x["Energy_kcal"])

        return {
            "Total_Generated": len(conformers),
            "Energy_Range": round(max(energies) - min_e, 2) if energies else 0,
            "Lowest_Energy": conformers[0] if conformers else None,
            "Conformers": conformers,
        }

    except Exception as e:
        return {"Error": str(e), "Conformers": []}
