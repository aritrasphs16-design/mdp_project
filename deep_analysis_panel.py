"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   CHEMOFILTER — DEEP ANALYSIS PANEL & INTERACTIVE TESTING                   ║
║   3D Viewer · Functional Groups · Reactivity · Metabolism · Synthesis       ║
║   Filtering · Sorting · Comparison · Bookmarking · Expandable Rows          ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""


# ═══════════════════════════════════════════════════════════════════
# DEEP ANALYSIS DATA BUILDER
# ═══════════════════════════════════════════════════════════════════

def build_deep_analysis(mol, compound_data, extended_data):
    """Build comprehensive deep analysis data for a compound."""
    from rdkit import Chem
    from rdkit.Chem import Descriptors, AllChem
    from rdkit.Chem.rdMolDescriptors import CalcNumRings, CalcNumAromaticRings
    import hashlib

    smiles = Chem.MolToSmiles(mol)
    seed = int(hashlib.md5(smiles.encode()).hexdigest(), 16) % 10**8

    # Functional Group Breakdown
    fg_patterns = {
        "Hydroxyl (-OH)": "[OX2H1]",
        "Carboxylic Acid (-COOH)": "[CX3](=O)[OX2H1]",
        "Primary Amine (-NH₂)": "[NX3;H2;!$(N-C=O)]",
        "Secondary Amine (>NH)": "[NX3;H1;!$(N-C=O)]",
        "Tertiary Amine (>N-)": "[NX3;H0;!$(N-C=O);!$(N=*)]",
        "Amide (-CONH-)": "[NX3][CX3](=[OX1])",
        "Ester (-COO-)": "[CX3](=O)[OX2H0]",
        "Aldehyde (-CHO)": "[CX3H1](=O)",
        "Ketone (>C=O)": "[CX3](=O)([#6])[#6]",
        "Alkyl Halide (-C-X)": "[CX4][F,Cl,Br,I]",
        "Aryl Halide (Ar-X)": "c[F,Cl,Br,I]",
        "Nitro (-NO₂)": "[N+](=O)[O-]",
        "Nitrile (-CN)": "[CX2]#N",
        "Sulfonamide (-SO₂NH-)": "[NX3][SX4](=O)(=O)",
        "Thiol (-SH)": "[SX2H]",
        "Thioether (-S-)": "[SX2]([#6])[#6]",
        "Sulfone (-SO₂-)": "[SX4](=O)(=O)",
        "Phosphate (-PO₄)": "[PX4](=O)(O)(O)",
        "Epoxide": "C1OC1",
        "Alkene (C=C)": "C=C",
        "Alkyne (C≡C)": "C#C",
    }

    functional_groups = []
    for name, smarts in fg_patterns.items():
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            count = len(mol.GetSubstructMatches(pat))
            functional_groups.append({"Group": name, "Count": count})

    # Reactivity Hotspots
    reactivity_hotspots = []
    hotspot_checks = [
        ("[CX3H1](=O)", "Electrophilic", "Aldehyde — nucleophilic addition"),
        ("[NX3;H2]", "Nucleophilic", "Free amine — electrophile target"),
        ("C=C", "Radical/Electrophilic", "Alkene — addition reactions"),
        ("[OX2H1]", "Nucleophilic", "Hydroxyl — oxidation/substitution"),
        ("[SX2H]", "Nucleophilic", "Thiol — strong nucleophile"),
        ("[F,Cl,Br,I][CX4]", "Electrophilic", "Alkyl halide — SN2/E2"),
        ("C1OC1", "Electrophilic", "Epoxide — ring-opening"),
    ]

    for smarts, nature, description in hotspot_checks:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            matches = mol.GetSubstructMatches(pat)
            reactivity_hotspots.append({
                "Nature": nature,
                "Description": description,
                "Sites": len(matches),
                "Atom_Indices": [m[0] for m in matches[:5]],
            })

    # Predicted Metabolism Sites
    metabolism_sites = []
    metab_checks = [
        ("[N;H0;X3;!$(N-C=O)][CH3,CH2,CH]", "N-Dealkylation", "CYP3A4/CYP2D6"),
        ("[O;H0;X2;!$(O-C=O)][CH3,CH2,CH]", "O-Dealkylation", "CYP2C19/CYP3A4"),
        ("[S;X2]", "S-Oxidation", "FMO/CYP3A4"),
        ("c1ccccc1", "Aromatic Hydroxylation", "CYP1A2/CYP2C9"),
        ("[CX4;H3,H2]", "Aliphatic Hydroxylation", "CYP3A4"),
        ("C=C", "Epoxidation", "CYP enzymes"),
        ("[OH]", "Glucuronidation", "UGT enzymes"),
    ]

    for smarts, reaction, enzyme in metab_checks:
        pat = Chem.MolFromSmarts(smarts)
        if pat and mol.HasSubstructMatch(pat):
            metabolism_sites.append({
                "Reaction": reaction,
                "Responsible_Enzyme": enzyme,
                "Sites": len(mol.GetSubstructMatches(pat)),
                "Phase": "Phase I" if "CYP" in enzyme else "Phase II",
            })

    # Synthetic Pathway Suggestions
    synthetic_suggestions = []
    if mol.HasSubstructMatch(Chem.MolFromSmarts("c1ccccc1[NX3]")):
        synthetic_suggestions.append("Buchwald-Hartwig amination for C-N bond formation")
    if mol.HasSubstructMatch(Chem.MolFromSmarts("c1ccccc1c2ccccc2")):
        synthetic_suggestions.append("Suzuki coupling for biaryl bond")
    if mol.HasSubstructMatch(Chem.MolFromSmarts("C(=O)[NX3]")):
        synthetic_suggestions.append("Amide coupling (EDC/HOBt or T3P)")
    if mol.HasSubstructMatch(Chem.MolFromSmarts("[OH]c1ccccc1")):
        synthetic_suggestions.append("O-Alkylation with Mitsunobu or SN2")
    if CalcNumRings(mol) >= 2:
        synthetic_suggestions.append("Ring construction via Diels-Alder or RCM")
    if not synthetic_suggestions:
        synthetic_suggestions.append("Standard linear synthesis recommended")

    return {
        "Functional_Groups": functional_groups,
        "Reactivity_Hotspots": reactivity_hotspots,
        "Metabolism_Sites": metabolism_sites,
        "Synthetic_Suggestions": synthetic_suggestions,
        "Total_FG": len(functional_groups),
        "Total_Hotspots": len(reactivity_hotspots),
        "Total_Metab_Sites": sum(s["Sites"] for s in metabolism_sites),
    }


# ═══════════════════════════════════════════════════════════════════
# INTERACTIVE TESTING FEATURES
# ═══════════════════════════════════════════════════════════════════

def filter_compounds(compounds, mw_range=None, logp_range=None, tpsa_range=None,
                     min_qed=None, only_lipinski=False, only_veber=False):
    """Filter compounds by property ranges."""
    filtered = list(compounds)

    if mw_range:
        filtered = [c for c in filtered if mw_range[0] <= c.get("MW", 0) <= mw_range[1]]
    if logp_range:
        filtered = [c for c in filtered if logp_range[0] <= c.get("LogP", 0) <= logp_range[1]]
    if tpsa_range:
        filtered = [c for c in filtered if tpsa_range[0] <= c.get("tPSA", 0) <= tpsa_range[1]]
    if min_qed is not None:
        filtered = [c for c in filtered if c.get("QED", 0) >= min_qed]
    if only_lipinski:
        filtered = [c for c in filtered if c.get("_vc", 5) == 0]
    if only_veber:
        filtered = [c for c in filtered if c.get("_rot", 99) <= 10 and c.get("_tp", 999) <= 140]

    return filtered


def sort_compounds(compounds, sort_by="LeadScore", ascending=False):
    """Sort compounds by any numeric property."""
    try:
        return sorted(compounds, key=lambda c: float(c.get(sort_by, 0)), reverse=not ascending)
    except:
        return compounds


def highlight_lipinski_violations(compound):
    """Return list of Lipinski violations for highlighting."""
    violations = []
    if compound.get("MW", 0) > 500:
        violations.append(("MW", f"{compound['MW']} > 500 Da"))
    if compound.get("LogP", 0) > 5:
        violations.append(("LogP", f"{compound['LogP']} > 5"))
    if compound.get("HBD", 0) > 5:
        violations.append(("HBD", f"{compound['HBD']} > 5"))
    if compound.get("HBA", 0) > 10:
        violations.append(("HBA", f"{compound['HBA']} > 10"))
    return violations


def compare_compounds_data(compounds):
    """Prepare data for side-by-side compound comparison."""
    if len(compounds) < 2:
        return []

    properties = ["MW", "LogP", "tPSA", "QED", "HBD", "HBA", "RotBonds",
                   "LeadScore", "SA_Score", "ArRings", "Fsp3"]

    comparison = []
    for prop in properties:
        row = {"Property": prop}
        for i, c in enumerate(compounds):
            row[c.get("ID", f"Cpd-{i+1}")] = c.get(prop, "N/A")
        comparison.append(row)

    return comparison


def get_compound_ranking(compounds, metric="LeadScore"):
    """Rank compounds by a given metric."""
    try:
        sorted_cpds = sorted(compounds, key=lambda c: float(c.get(metric, 0)), reverse=True)
    except:
        sorted_cpds = compounds

    rankings = []
    for i, c in enumerate(sorted_cpds):
        rankings.append({
            "Rank": i + 1,
            "ID": c.get("ID", "Unknown"),
            "Value": c.get(metric, "N/A"),
            "Grade": c.get("Grade", "N/A"),
            "Medal": "🥇" if i == 0 else ("🥈" if i == 1 else ("🥉" if i == 2 else "")),
        })

    return rankings


# ═══════════════════════════════════════════════════════════════════
# EXPANDED ROW DETAILS HTML BUILDER
# ═══════════════════════════════════════════════════════════════════

def build_expanded_row_html(compound, extended):
    """Build HTML for expandable row detail view."""
    badge = extended.get("Drug_Likeness_Badge", "Unknown")
    badge_color = extended.get("Badge_Color", "#666")
    warnings = extended.get("Property_Warnings", [])

    # Warning pills
    warning_html = ""
    for w in warnings:
        warning_html += f'<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:0.6rem;background:rgba(248,113,113,0.1);color:#f87171;border:1px solid rgba(248,113,113,0.2);margin:2px">{w}</span>'

    if not warning_html:
        warning_html = '<span style="color:#4ade80;font-size:0.6rem">✓ No warnings</span>'

    html = f"""
    <div style="padding:16px;background:rgba(5,8,15,0.6);border:1px solid rgba(232,160,32,0.08);border-radius:8px;margin:8px 0">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
            <span style="font-family:'DM Serif Display',serif;font-size:1.1rem;color:white">{compound.get('ID','Unknown')}</span>
            <span style="padding:4px 12px;border-radius:20px;font-size:0.65rem;font-weight:700;background:{badge_color}22;color:{badge_color};border:1px solid {badge_color}44">{badge}</span>
        </div>

        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;font-size:0.7rem;font-family:'IBM Plex Mono',monospace">
            <div>
                <div style="color:rgba(200,222,255,0.4);margin-bottom:4px">ADMET</div>
                <div style="color:#c8deff">HIA: {extended.get('HIA','N/A')}</div>
                <div style="color:#c8deff">BBB: {extended.get('BBB_Penetration','N/A')}</div>
                <div style="color:#c8deff">CYP Risk: {extended.get('CYP450_Risk','N/A')}</div>
                <div style="color:#c8deff">Solubility: {extended.get('Solubility_Class','N/A')}</div>
            </div>
            <div>
                <div style="color:rgba(200,222,255,0.4);margin-bottom:4px">STRUCTURE</div>
                <div style="color:#c8deff">Chiral: {extended.get('Chiral_Centers',0)}</div>
                <div style="color:#c8deff">Complexity: {extended.get('Topological_Complexity','N/A')}</div>
                <div style="color:#c8deff">Symmetry: {extended.get('Molecular_Symmetry','N/A')}%</div>
                <div style="color:#c8deff">Fragment Div: {extended.get('Fragment_Diversity','N/A')}</div>
            </div>
            <div>
                <div style="color:rgba(200,222,255,0.4);margin-bottom:4px">OPTIMIZATION</div>
                <div style="color:#c8deff">Lead Potential: {extended.get('Lead_Optimization_Potential','N/A')}/100</div>
                <div style="color:#c8deff">Lig Efficiency: {extended.get('Ligand_Efficiency','N/A')}</div>
                <div style="color:#c8deff">LipE: {extended.get('Lipophilic_Efficiency','N/A')}</div>
                <div style="color:#c8deff">Synth: {extended.get('Synthetic_Difficulty','N/A')}</div>
            </div>
        </div>

        <div style="margin-top:12px">{warning_html}</div>
    </div>
    """
    return html
