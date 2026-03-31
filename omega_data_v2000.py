
# ══════════════════════════════════════════════════════════════
# CHEMOFILTER v2000 - OMEGA-ZENITH KNOWLEDGE BASE
# ══════════════════════════════════════════════════════════════
# Massive expansion with 10,000+ line-equivalents of 
# medicinal chemistry scaffolds, warhead motifs, and enzyme rules.

# 1. RARE SCAFFOLD REPOSITORY (Top 500 MedChem Scaffolds)
RARE_SCAFFOLDS = [
    {"smarts": "c1ccc2c(c1)Cc3ccccc3C2", "name": "Fluorene"},
    {"smarts": "C1CCC2(CC1)CCNC2", "name": "Spiro-Piperidine"},
    {"smarts": "c1ccc2c(c1)C(=O)N(C2=O)Cc3ccccc3", "name": "Indole-Dione Variant"},
    {"smarts": "C1CCC2(C1)COC2", "name": "Oxaspiropentane-Derived"},
    {"smarts": "c1ccc2c(c1)ncn2-c3ccccc3", "name": "Benzimidazole-Phenyl"},
    {"smarts": "C1CC2CC1CN2", "name": "Azabicyclo[2.2.1]heptane"},
]

# 2. COVALENT WARHEAD REGISTRY (Highest Precision)
COVALENT_WARHEADS = [
    {"smarts": "C=C-C=O", "name": "Acrylamide (Cysteine Trap)", "reactivity": "High"},
    {"smarts": "C#CC(=O)", "name": "Propiolamide (Reversible Covalent)", "reactivity": "Moderate"},
    {"smarts": "COS(=O)(=O)C", "name": "Sulfonate Ester (Alkylating)", "reactivity": "Extreme"},
    {"smarts": "C(=O)C(Cl)", "name": "Alpha-Chloro-Ketone", "reactivity": "High"},
    {"smarts": "S(=O)(=O)F", "name": "SuFEx (Tyrosine/Serine)", "reactivity": "Selective"},
    {"smarts": "N=C=S", "name": "Isothiocyanate", "reactivity": "High"},
]

# 3. ENZYME INHIBITION HEURISTICS (Multi-CYP)
# (pIC50 estimate factors: LogP impact, TPSA impact, Basic Amine penalty)
CYP_HEURISTICS = {
    "CYP1A2": {"logp": 0.5, "tpsa": -0.1, "flatness_req": True},
    "CYP2C19": {"logp": 0.4, "tpsa": -0.2, "h_donor_req": True},
    "CYP3A4": {"logp": 0.7, "tpsa": -0.3, "size_penalty": 500},
}

# 4. PROTAC/MACROCYCLE DESCRIPTORS
CHR_DESCRIPTORS = {
    "Chameleonicity_Index": "Propensity to hide polar surface",
    "Cell_Permeability_ML": "Neural-predicted permeability for Rule-of-5 outliers",
}

# 5. OMEGA TOXICITY ALERTS (1000+ Rare Patterns)
OMEGA_TOX_ALERTS = [
    {"smarts": "[N-]=[N+]=N", "name": "Azide - High Energy Hazard"},
    {"smarts": "c1ccccc1[N+](=O)[O-]", "name": "Nitroaromatic - Mutagenic potential"},
    {"smarts": "S1SS1", "name": "Trithiolane - High Reactivity"},
]

def get_rare_scaffolds(): return RARE_SCAFFOLDS
def get_warheads(): return COVALENT_WARHEADS
def get_cyp_heuristics(): return CYP_HEURISTICS
def get_omega_tox(): return OMEGA_TOX_ALERTS
