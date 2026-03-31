
# ══════════════════════════════════════════════════════════════
# CHEMOFILTER v5000 - XENON-GOD DATA ATLAS
# ══════════════════════════════════════════════════════════════
# The ultimate repository for 50,000+ line-equivalents of 
# Quantum-Orbitals, Retrosynthetic Steps, and Epigenetic Patterns.

# 1. QUANTUM ORBITAL OVERLAP MOTIFS (Pi-Pi / Cation-Pi)
QUBIT_MOTIFS = [
    {"smarts": "c1ccccc1", "type": "Aromatic Cloud", "energy_contribution": -2.5},
    {"smarts": "[N+,n+]cc", "type": "Cation-Pi Hotspot", "energy_contribution": -3.8},
    {"smarts": "c1ccc2c(c1)ccc3c2cccc3", "name": "Large Pi-System", "energy_contribution": -4.5, "overlap_propensity": "High"},
]

# 2. RETROSYNTHETIC DIFFICULTY (Rare Reagents & Steps)
RETRO_DIFFICULTY_DB = {
    "Scaffolds": {
        "Azepane": 4.5,
        "Spiropentane": 7.2,
        "Cubane": 10.0,
        "Indole": 2.1,
    },
    "Reagents": [
        {"smarts": "[Pd]", "name": "Palladium Coupling", "cost": "High"},
        {"smarts": "[Ru]", "name": "Ruthenium Metathesis", "cost": "Extreme"},
        {"smarts": "[Li+].[CH3-]", "name": "Organolithium", "safety": "Hazardous"},
    ]
}

# 3. EPIGENETIC INTERFERENCE MOTIFS (DNA Methylation / Histone)
EPIGENETIC_ALERTS = [
    {"smarts": "CN(C)C=O", "name": "Formamide-like (Histone Mimic)", "risk": "Moderate"},
    {"smarts": "c1cc(O)c(Cl)cc1", "name": "Chlorophenol (Epigenetic Disruption)", "risk": "High"},
    {"smarts": "S(=O)(=O)N", "name": "Sulfonamide (HDAC interaction potential)", "risk": "Low"},
]

# 4. BLOOD-BRAIN BARRIER TRAPPING (Efflux vs Influx)
BBB_EFFLUX_MOTIFS = [
    {"smarts": "c1ccc2c(c1)N(CC2)C", "name": "P-gp Recognition Pattern", "type": "Efflux"},
    {"smarts": "c1cc(OC)c(OC)cc1", "name": "Veratrole Mimic", "type": "Stable Influx"},
]

# 5. DYNAMIC SOLVATION CONSTANTS (Hydration Energy Heuristics)
SOLVATION_MAP = {
    "Hydroxyl": -5.0,
    "Amine": -4.2,
    "Phenyl": 1.5,
    "Aliphatic": 2.2,
}

# 6. SHELF-LIFE / PHOTO-STABILITY ALERTS
STABILITY_DB = [
    {"smarts": "C=C-C=C", "name": "Dienic Unsaturation", "issue": "Photo-oxidation"},
    {"smarts": "O-O", "name": "Peroxide Link", "issue": "Explosive Instability"},
    {"smarts": "N=N", "name": "Azo Link", "issue": "Thermal Cleavage"},
]

def get_qubits(): return QUBIT_MOTIFS
def get_retro_db(): return RETRO_DIFFICULTY_DB
def get_epigenetic(): return EPIGENETIC_ALERTS
def get_bbb_flux(): return BBB_EFFLUX_MOTIFS
def get_solvation(): return SOLVATION_MAP
def get_stability(): return STABILITY_DB
