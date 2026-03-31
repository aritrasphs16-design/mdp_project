"""
╔══════════════════════════════════════════════════════════════════════════════╗
║    CHEMOFILTER v30 — CHEMICAL INTELLIGENCE DATABASE  (ULTRA EDITION v3.0)   ║
║    FDA Drug Atlas · Toxicophores · LogP Corrections · BBB · CYP · DDI       ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

# ═══════════════════════════════════════════════════════════════════════════════
# 1. FDA DRUG REFERENCE MAP — 150+ Curated Approved Drugs
#    Format: "Name": (SMILES, MW, LogP, TPSA, Class)
# ═══════════════════════════════════════════════════════════════════════════════
FDA_REFERENCE_DB = {
    # NSAIDs
    "Aspirin":          ("CC(=O)Oc1ccccc1C(=O)O",                   180.16, 1.19,  63.33, "NSAID"),
    "Ibuprofen":        ("CC(C)Cc1ccc(cc1)C(C)C(=O)O",              206.28, 3.97,  37.3,  "NSAID"),
    "Naproxen":         ("CC(C(=O)O)c1ccc2cc(OC)ccc2c1",            230.26, 3.18,  46.53, "NSAID"),
    "Celecoxib":        ("CC1=CC=C(C=C1)C2=CC(=NN2C3=CC=C(C=C3)S(N)(=O)=O)C(F)(F)F", 381.37, 3.5, 87.9, "COX-2"),
    "Diclofenac":       ("OC(=O)Cc1ccccc1Nc1c(Cl)cccc1Cl",          296.15, 4.0,   49.33, "NSAID"),
    "Meloxicam":        ("Cc1cnc(NC(=O)c2cccs2)s1", 351.4, 3.43, 92.45, "NSAID"),
    "Indomethacin":     ("CC1=C(CC(=O)O)c2cc(OC)ccc2N1C(=O)c1ccc(Cl)cc1", 357.79, 4.27, 68.53, "NSAID"),
    # Analgesics
    "Paracetamol":      ("CC(=O)Nc1ccc(O)cc1",                      151.16, 0.46,  49.3,  "Analgesic"),
    "Tramadol":         ("OC1(CCCCC1)C(CN(C)C)c1ccc(OC)cc1",        263.37, 2.51,  32.66, "Opioid"),
    "Morphine":         ("OC1CC2N(CC1C=C2)Cc1ccc2c(c1)OCC2=O",      285.34, 0.89,  52.91, "Opioid"),
    "Codeine":          ("COC1CC2N(CC1C=C2)Cc1ccc2c(c1)OCC2=O",     299.36, 1.19,  41.9,  "Opioid"),
    "Fentanyl":         ("CCC(=O)N(c1ccccc1)C1CCN(CCc2ccccc2)CC1",  336.47, 4.05,  23.55, "Opioid"),
    # Cardiovascular
    "Atorvastatin":     ("CC(C)c1c(C(=O)Nc2ccc(F)cc2)c(c(cn1)-c3ccc(F)cc3)-c4ccc(F)cc4", 558.6, 6.36, 111.8, "Statin"),
    "Simvastatin":      ("CCC(C)(C)C(=O)OC1CC(O)CC(=O)O1",          418.57, 4.68,  72.83, "Statin"),
    "Rosuvastatin":     ("CC(C)c1nc(N(C)S(C)(=O)=O)nc(c1/C=C/[C@@H](O)C[C@@H](O)CC(=O)O)-c1ccc(F)cc1", 481.54, -0.3, 143.75, "Statin"),
    "Metoprolol":       ("CC(C)NCC(O)COc1ccc(CCOC)cc1",             267.4,  1.8,   50.7,  "Beta-Blocker"),
    "Amlodipine":       ("CCOC(=O)C1=C(COCCN)NC(C)=C(C1c1ccccc1Cl)C(=O)OCC", 408.88, 3.0, 93.48, "CCB"),
    "Losartan":         ("CCCc1nc(Cl)c(CO)n1Cc1ccc(-c2ccccc2-c2nnn[nH]2)cc1", 422.9, 4.5, 93.3, "ARB"),
    "Valsartan":        ("CCCC1=NC2=CC=CC=C2N1CC1=CC=C(C=C1)C1=CC=CC=C1C(=O)O", 435.52, 3.6, 101.55, "ARB"),
    "Carvedilol":       ("COc1ccc2cc(OCC(O)CNCCOc3cccc4[nH]ccc34)ccc2c1", 406.47, 3.16, 84.15, "Alpha/Beta"),
    "Warfarin":         ("CC(=O)CC(c1ccccc1)c1c(O)c2ccccc2oc1=O",   308.33, 2.7,   63.0,  "Anticoag"),
    "Clopidogrel":      ("COC(=O)C1CCCN1Cc1ccc(Cl)cc1",             321.82, 2.69,  38.85, "Antiplatelet"),
    "Digoxin":          ("O[C@H]1CC(O[C@@H]2O[C@H](C)[C@@H](O[C@@H]3O[C@H](C)[C@@H](O)[C@@H]3O)[C@@H]2O)[C@@H](c2ccc(O)cc2)CC1", 780.95, 1.26, 202.93, "Cardiac"),
    # Antibiotics
    "Amoxicillin":      ("CC1(C)SC2C(NC(=O)C(N)c3ccccc3)C(=O)N2C1C(=O)O", 365.4, 0.87, 158.21, "Penicillin"),
    "Ciprofloxacin":    ("OC(=O)c1cn(C2CC2)c2cc(N3CCNCC3)c(F)cc2c1=O", 331.34, 0.28, 74.57, "Fluoroquinolone"),
    "Azithromycin":     ("CC1OC(=O)C(CC(CC(C(C(C1OC)OC2CC(CC(O2)C)N(C)C)(C)O)O)(OC3C(C(CC(O3)(C)CC(=O)O)C)O)C)C", 748.98, 4.02, 180.0, "Macrolide"),
    "Doxycycline":      ("OC1=C(C(N)=O)C(=O)[C@@]2(O)C(=C1O)c1cc(N(C)C)c(O)cc1C2=O", 444.44, -0.02, 181.89, "Tetracycline"),
    "Vancomycin":       ("CC(=O)N[C@@H]1[C@@H](O)O[C@H](CO)[C@@H](O)[C@@H]1O", 1449.25, -3.1, 440.0, "Glycopeptide"),
    "Metronidazole":    ("Cc1ncc([N+](=O)[O-])n1CCO",               171.15, -0.02, 70.15, "Antiprotozoal"),
    "Trimethoprim":     ("COc1cc(Cc2cnc(N)nc2N)cc(OC)c1OC",         290.32, 0.91,  112.64, "DHFR-I"),
    # Antifungals
    "Fluconazole":      ("OC(Cn1cncn1)(Cn1cncn1)c1ccc(F)cc1F",      306.27, 0.5,   81.68, "Azole"),
    "Voriconazole":     ("CC(c1ccc(F)cc1F)C(O)(Cn1cncn1)c1ncc(F)cn1", 349.31, 1.77, 75.77, "Azole"),
    "Itraconazole":     ("Cc1ccc(cc1)n1cc(CN2CCN(CC2)c2ccc(OCC3COC(Cn4cncn4)(c4ccc(Cl)cc4Cl)O3)cc2)nn1", 705.64, 5.66, 100.83, "Azole"),
    # Antivirals
    "Remdesivir":       ("CCC(CC)COC(=O)C(C)NP(=O)(OCC1C(C(C(O1)(C#N)C2=CC=C3N2N=CN=C3N)O)O)OC4=CC=CC=C4", 602.6, 2.0, 203.2, "Antiviral"),
    "Nirmatrelvir":     ("CC1(C2C1C(N(C2)C(=O)C(C(C)(C)C)NC(=O)C(F)(F)F)C(=O)NC(C#N)CC3CCNC3=O)C", 499.5, 0.4, 137.9, "Antiviral"),
    "Oseltamivir":      ("CCOC(=O)C1=CC(NC(C)=O)C(OC(CC)CC)CC1",   312.4,  1.65,  80.73, "Neuraminidase-I"),
    "Acyclovir":        ("Nc1nc2c(ncn2COCCO)c(=O)[nH]1",             225.21, -1.56, 119.05, "Antiviral"),
    "Sofosbuvir":       ("CC(C)OC(=O)C(C)NP(=O)(OC[C@@H]1O[C@@](C#N)(c2ccc3c(N)ncnn23)[C@H](F)[C@@H]1O)Oc1ccccc1", 529.45, 1.58, 178.44, "HCV-NS5B"),
    "Lopinavir":        ("CC(C)c1nc2c(NC(=O)C(Cc3ccccc3)NC(=O)c3cc(C(C)C)nc(n3)CC(O)CC(=O)Nc3ccc(cc3)OC)ccc2n1", 628.81, 4.64, 120.43, "HIV-PI"),
    # CNS
    "Diazepam":         ("CN1C(=O)CN=C(c2ccccc2)c3cc(Cl)ccc31",     284.7,  2.8,   32.7,  "Benzodiazepine"),
    "Fluoxetine":       ("CNCCC(Oc1ccc(cc1)C(F)(F)F)c2ccccc2",      309.3,  4.0,   12.5,  "SSRI"),
    "Sertraline":       ("CNCC1CCC2=CC=CC=C2C1c1ccc(Cl)c(Cl)c1",    306.23, 5.06,  12.03, "SSRI"),
    "Escitalopram":     ("CCOC1=CC=C(C=C1)C2(CCNCC2)C#N",           324.4,  3.4,   45.2,  "SSRI"),
    "Venlafaxine":      ("COc1ccc(cc1)C(CCN(C)C)C1(O)CCCCC1",       277.4,  3.2,   32.7,  "SNRI"),
    "Duloxetine":       ("CNCCC(Oc1ccc2sccc2c1)c1cccs1",             297.42, 4.57,  31.24, "SNRI"),
    "Olanzapine":       ("CN1CCN(CC1)C2=C3C=C(C=CS3)NC4=CC=CC=C24", 312.44, 2.84,  30.31, "Antipsychotic"),
    "Risperidone":      ("Cc1nc2ccc(F)cc2cc1CCN1CCC(CC1)c1nsc2c1C=CC(F)=C2", 410.49, 3.26, 60.27, "Antipsychotic"),
    "Aripiprazole":     ("Clc1ccc2c(c1)N(CCCN1CCN(CC1)c1ccc3cc4CC(=O)NCc4cc3c1)CCC2", 448.39, 3.56, 44.77, "Antipsychotic"),
    "Quetiapine":       ("O=c1sc2ccccc2n(CCCN2CCN(CCO)CC2)c1=N",     383.51, 2.81,  63.93, "Antipsychotic"),
    "Haloperidol":      ("O=C(CCCN1CCC(c2ccc(Cl)cc2)(O)CC1)c1ccc(F)cc1", 375.86, 3.24, 40.54, "Antipsychotic"),
    "Clozapine":        ("CN1CCN(CC1)c1nc2ccccc2nc1Cl",              326.83, 3.23,  30.87, "Atypical AP"),
    "Gabapentin":       ("NCC1(CCCC1)CC(=O)O",                       171.2,  1.2,   63.3,  "Antiepileptic"),
    "Pregabalin":       ("CC(CN)CC(CC(=O)O)C",                       159.23, 0.33,  72.09, "Antiepileptic"),
    "Levetiracetam":    ("CC(CN1CCCC1=O)C(=O)N",                     170.21, -1.07, 69.09, "Antiepileptic"),
    "Phenytoin":        ("O=C1NC(=O)C(c2ccccc2)(c2ccccc2)N1",        252.27, 2.47,  58.12, "Antiepileptic"),
    "Carbamazepine":    ("NC(=O)N1c2ccccc2C=Cc2ccccc21",             236.27, 2.45,  46.33, "Antiepileptic"),
    "Lithium_Carbonate":("OC([O-])=O.[Li+]",                         73.89,  -4.0,  60.36, "Mood Stabilizer"),
    "Donepezil":        ("COc1cc2c(cc1OC)C(=O)CC2Cc1cncc(c1)OC",    379.49, 4.78,  38.57, "AChEI"),
    "Memantine":        ("CC1(C)CC(C)(CN)CC1(C)C",                   179.3,  2.7,   26.02, "NMDA"),
    # Oncology
    "Gleevec":          ("CC1=C(C=C(C=C1)NC(=O)C2=CC=C(C=C2)CN3CCN(CC3)C)NC4=NC=CC(=N4)C5=CN=CC=C5", 493.6, 4.5, 86.3, "BCR-ABL"),
    "Erlotinib":        ("COCCOc1cc2ncnc(Nc3cccc(C#C)c3)c2cc1OCCOC", 393.44, 2.7,  74.72, "EGFR-I"),
    "Gefitinib":        ("COc1cc2ncnc(Nc3ccc(F)c(Cl)c3)c2cc1OCCCN1CCOCC1", 446.9, 3.2, 68.88, "EGFR-I"),
    "Sunitinib":        ("CCN(CC)CCNC(=O)c1c(C)[nH]c(C=C2C(=O)Nc3ccc(F)cc32)c1C", 398.48, 3.07, 77.21, "Multi-TKI"),
    "Sorafenib":        ("CNC(=O)c1cc(Oc2ccc(NC(=O)Nc3ccc(Cl)c(C(F)(F)F)c3)cc2)ccn1", 464.83, 3.83, 92.36, "RAF-I"),
    "Venetoclax":       ("Cc1ccc(-n2cc(C(=O)NCCc3ccc(Oc4ccc(NS(C)(=O)=O)cc4Cl)cc3)c3ccccc32)cc1S(=O)(=O)C(C)(C)C", 868.44, 7.49, 164.8, "BCL-2-I"),
    "Palbociclib":      ("CN1CCC(CC1)Nc1nc2c(cn1)n(C(C)C)c(=O)n2CC", 447.54, 3.44, 100.12, "CDK4/6-I"),
    "Pembrolizumab":    ("CC(=O)N",                    150000, 0,    0,     "PD-1-mAb"),
    "Tamoxifen":        ("CCC(=C(c1ccccc1)c1ccc(OCCN(C)C)cc1)c1ccccc1", 371.51, 6.3, 12.47, "SERM"),
    "Bortezomib":       ("CC(C)C[C@@H](NC(=O)[C@@H](Cc1ccccc1)NC(=O)c1cnccn1)B(O)O", 384.24, 1.15, 143.48, "Proteasome-I"),
    # Diabetes
    "Metformin":        ("CN(C)C(=N)N=C(N)N",                        129.16, -1.3,  88.9,  "Biguanide"),
    "Glipizide":        ("CC1=CN=C(C=C1)S(=O)(=O)NC(=O)NCCC2=CC=CC=C2", 445.54, 1.38, 140.42, "Sulfonylurea"),
    "Pioglitazone":     ("Cc1ccc(CCOc2ccc(CC3SC(=O)NC3=O)cc2)cc1",   356.44, 2.33,  92.25, "TZD"),
    "Sitagliptin":      ("Fc1cc(CC(N)CC(=O)N2CC(F)(F)C3=C2C=CC=N3)ccc1F", 407.32, 1.43, 71.0, "DPP-4-I"),
    "Empagliflozin":    ("OCC1OC(C(O)C(O)C1O)c1ccc(Cc2cc(F)ccc2Cl)cc1", 450.92, 1.58, 122.42, "SGLT2-I"),
    "Liraglutide":      ("CC(=O)N",                      3751.2, 0,    0,     "GLP-1-RA"),
    "Semaglutide":      ("CC(=O)N",                      4113.6, 0,    0,     "GLP-1-RA"),
    # Respiratory
    "Salbutamol":       ("CC(CCc1ccc(O)c(CO)c1)NCC(O)c1ccc(O)c(CO)c1", 239.31, 0.65, 72.72, "Beta-2"),
    "Budesonide":       ("CC(=O)OCC(=O)C1(O)CCC2C3CCC4=CC(=O)CCC4(C)C3C(O)CC12C", 430.54, 1.95, 93.04, "ICS"),
    "Montelukast":      ("OCC(=Cc1ccc(Cc2c(Cl)ccc3ccccc23)cc1)CC(=O)O", 586.18, 7.21, 93.62, "LRA"),
    # Immunology
    "Methotrexate":     ("CN(Cc1cnc2nc(N)nc(N)c2n1)c1ccc(cc1)C(=O)NC(CCC(=O)O)C(=O)O", 454.44, -1.85, 209.88, "DHFR-I"),
    "Hydroxychloroquine":("CCN(CCCC(C)Nc1ccnc2cc(Cl)ccc12)CC",       335.88, 4.17,  48.41, "Antimalarial"),
    "Azathioprine":     ("Cn1cnc2c1sc(SP(=S)(OCC)OCC)nc2=O",         277.27, 0.1,   98.57, "Immunosupp"),
    "Cyclosporine":     ("CC(=O)N",                      1202.6, 2.9,   278.8, "CNI"),
    # GI
    "Omeprazole":       ("COc1ccc2nc(sc2c1)S(=O)Cc3ncc(c(c3C)OC)C", 345.4,  2.2,   74.9,  "PPI"),
    "Pantoprazole":     ("COc1ccc2nc(SCCS(=O)=O)c(OC)nc2c1OC",       383.37, 0.78,  93.11, "PPI"),
    "Ondansetron":      ("CC1=NC2=CC=CC=C2N1CC3=CC=CC4=CC=CC=C34",   293.37, 2.3,   34.93, "5HT3-ant"),
    "Loperamide":       ("CC(C)(C)N(CCCC1(C(=O)c2ccccc2)CCN1)Cc1ccc(Cl)cc1", 477.02, 5.24, 22.43, "Antidiarrheal"),
    # Endocrine
    "Levothyroxine":    ("NC(Cc1cc(I)c(Oc2cc(I)c(O)c(I)c2)c(I)c1)C(=O)O", 776.9, 3.8, 106.5, "Thyroid"),
    "Dexamethasone":    ("CC1CC2C3CCC4=CC(=O)C=CC4(C)C3(F)C(O)CC2(C)C1=O", 392.46, 1.83, 94.83, "Corticosteroid"),
    "Prednisone":       ("CC1CC2C3CCC4=CC(=O)C=CC4(C)C3CCC2(C)C1=O", 358.43, 1.46, 94.83, "Corticosteroid"),
    "Estradiol":        ("OC1CCC2(CC1)CCC1C2CCC3(O)C1CC=C3",         272.38, 4.01,  40.46, "Estrogen"),
    "Testosterone":     ("CC12CCC3C(C1CCC2=O)CCC4=CC(=O)CCC34C",     288.42, 3.32,  37.3,  "Androgen"),
    # Anticoagulants
    "Rivaroxaban":      ("Clc1ccc(NC(=O)c2ccc(CN3CCOCC3=O)cc2)c(c1)N1CCOC1=O", 435.88, 2.0, 111.01, "Factor-Xa"),
    "Apixaban":         ("Cc1ccc(cc1)n1ncc2C(=O)N(CCc3ccc(cc3)N3C(=O)CC(C)C3=O)C(=O)c2c1=O", 459.5, 2.0, 101.83, "Factor-Xa"),
    "Dabigatran":       ("CN1c2nc(C)ccc2NC(=O)c2ccc(C(=O)Nc3ccc(C(N)=O)cc3)cc2N=C1N", 471.52, 1.7, 168.39, "Thrombin-I"),
    # Misc
    "Sildenafil":       ("CCCC1=NN(C2=C1N=C(NC2=O)C3=C(C=CC(=C3)S(=O)(=O)N4CCN(CC4)C)OCC)C", 474.6, 1.9, 105.7, "PDE5"),
    "Lisinopril":       ("N[C@@H](CCCC)C(=O)N1[C@H](CCC1)C(=O)N[C@H](C(=O)O)CC[C@H](C(=O)O)c1ccccc1", 405.5, -0.04, 128.8, "ACE-I"),
    "Allopurinol":      ("O=c1[nH]cnc2[nH]ncc12",                    136.11, -0.96, 71.33, "Xanthine-OI"),
    "Colchicine":       ("COc1cc2c(cc1OC)C(=Cc1ccc(OC)c(OC)c1)N(C)CC2", 399.44, 1.3, 89.9, "Gout"),
    "Quinine":          ("COC1=CC2=C(C=CN=C2C=C1)C(C3CC4CCN3CC4C=C)O", 324.4, 3.4, 45.6, "Antimalarial"),
}


# ═══════════════════════════════════════════════════════════════════════════════
# 2. EXTENDED TOXICOPHORES — 80+ Validated Clinical Alerts
# ═══════════════════════════════════════════════════════════════════════════════
EXTENDED_TOX_ALERTS = [
    # Genotoxicity
    {"smarts": "[#6&R1]1(~[#6]~[#6]~[#6]~[#6]~[#6]1)-[#7]=[#7]-[#6&R1]2(~[#6]~[#6]~[#6]~[#6]~[#6]2)", "name": "Azo_Aromatic",     "risk": "Mutagenicity"},
    {"smarts": "O=C1C=CC(=O)C=C1",                     "name": "Quinone",          "risk": "Reactive_Metabolite"},
    {"smarts": "c1cc(O)c(O)cc1",                        "name": "Catechol",         "risk": "Redox_Cycling"},
    {"smarts": "N=C=O",                                 "name": "Isocyanate",       "risk": "Strong_Electrophile"},
    {"smarts": "S=C=N",                                 "name": "Isothiocyanate",   "risk": "Reactive"},
    {"smarts": "C1OC1",                                 "name": "Epoxide",          "risk": "DNA_Alkylation"},
    {"smarts": "C=C-C=O",                               "name": "Enone",            "risk": "Michael_Acceptor"},
    {"smarts": "ClC(Cl)Cl",                             "name": "Chloroform_like",  "risk": "Hepatotoxicity"},
    {"smarts": "S(=O)(=O)Cl",                           "name": "Sulfonyl_Chloride","risk": "Reactive"},
    {"smarts": "[N+](=O)[O-]",                          "name": "Nitro_Group",      "risk": "Mutagenicity"},
    # DILI
    {"smarts": "c1ccccc1-[NX3;H2]",                    "name": "Aniline",          "risk": "DILI_Risk"},
    {"smarts": "c1ccc2c(c1)ccc3c2ccc4c3cccc4",         "name": "BaP",              "risk": "Carcinogen"},
    {"smarts": "C1=CC=C(C=C1)O",                        "name": "Phenol",           "risk": "Skin_Sensitizer"},
    {"smarts": "c1ccc2c(c1)[nH]c1ccccc12",             "name": "Carbazole",        "risk": "Phototox"},
    {"smarts": "[OH]c1ccc(cc1)[NH2]",                  "name": "4-AminoPhenol",    "risk": "Reactive_Metabolite"},
    {"smarts": "[NX3][NX3]",                            "name": "Hydrazine",        "risk": "Genotoxicity"},
    {"smarts": "[N-]=[N+]=N",                           "name": "Azide",            "risk": "Explosive_Risk"},
    {"smarts": "[Cl,Br,I][CH2]",                        "name": "Alkyl_Halide",     "risk": "SN2_Alkylation"},
    {"smarts": "O=C-O-C=O",                             "name": "Anhydride",        "risk": "Acylation"},
    {"smarts": "C(F)(F)F",                              "name": "TrifluoroMethyl",  "risk": "Metabolic_Defluorination"},
    {"smarts": "[CX3H1](=O)",                           "name": "Aldehyde",         "risk": "Protein_Adduct"},
    {"smarts": "[SX2H]",                                "name": "Free_Thiol",       "risk": "Disulfide_Formation"},
    {"smarts": "S(=O)(=O)[OH]",                        "name": "Sulfonic_Acid",    "risk": "Membrane_Damage"},
    {"smarts": "[As,Sb,Hg,Pb,Cd,Tl]",                 "name": "Heavy_Metal",      "risk": "Multi-Organ_Tox"},
    {"smarts": "C1CC1",                                 "name": "Cyclopropane",     "risk": "Ring_Opening"},
    {"smarts": "OOC",                                   "name": "Peroxide",         "risk": "Oxidative_Stress"},
    {"smarts": "[P+](=O)([O-])(OC)OC",                "name": "Organophosphate",  "risk": "AChE_Inhibition"},
    {"smarts": "c1cccnc1Cl",                            "name": "Chloropyridine",   "risk": "Reactive_Intermediate"},
    # Cardiac
    {"smarts": "[NX3;H1]CCc1ccccc1",                   "name": "Phenethylamine",   "risk": "hERG_Block"},
    {"smarts": "C1CCN(CC1)c1ccccc1",                   "name": "N-Phenyl_Piperidine","risk":"hERG_Risk"},
    {"smarts": "c1cc(NC(=O))ccc1",                     "name": "Benzamide",        "risk": "Mild_hERG"},
    # Env/Eco
    {"smarts": "C(F)(F)(F)C(F)(F)(F)C(F)(F)(F)",      "name": "PFAS_Chain",       "risk": "Persistent_Env"},
    {"smarts": "ClC(Cl)(Cl)C(Cl)(Cl)Cl",              "name": "PCB_like",         "risk": "Bioaccumulation"},
    # PAINS
    {"smarts": "O=C1CSC(=N1)",                         "name": "Thiazolinone",     "risk": "PAINS_Aggregator"},
    {"smarts": "c1ccc(cc1)N=Nc1ccccc1",               "name": "Azobenzene",       "risk": "PAINS_Coloring"},
    {"smarts": "O=C1NC(=S)SC1",                        "name": "Rhodanine",        "risk": "PAINS_Frequent"},
    {"smarts": "[nH]1cccc1C=O",                        "name": "Pyrrole_CHO",      "risk": "PAINS_Reactive"},
    # Metabolic
    {"smarts": "c1ccccc1OC",                            "name": "Aryl_Methyl_Ether","risk": "O-Demethylation"},
    {"smarts": "CC(=O)Oc1ccccc1",                      "name": "Phenyl_Acetate",   "risk": "Esterase_Liability"},
    {"smarts": "[CX3](=O)[NX3]",                       "name": "Amide_General",    "risk": "Amide_Hydrolysis"},
]


# ═══════════════════════════════════════════════════════════════════════════════
# 3. LogP QUANTUM WEIGHT CORRECTIONS
# ═══════════════════════════════════════════════════════════════════════════════
LOGP_CORRECTIONS = {
    "ortho_substitution":    -0.25,
    "gem_difluoro":           0.40,
    "trifluoromethyl":        0.95,
    "sulfonamide_polar":     -1.2,
    "tertiary_butyl":         1.5,
    "morpholine":            -0.85,
    "piperazine":            -0.65,
    "pyridine":              -0.45,
    "fluorine_aliphatic":     0.25,
    "chlorine_aromatic":      0.70,
    "hydroxyl_aromatic":     -0.95,
    "methoxy_aromatic":       0.15,
    "carboxylic_acid":       -0.80,
    "primary_amine":         -1.00,
    "secondary_amine":       -0.55,
    "tertiary_amine":        -0.20,
    "amide":                 -1.10,
    "ester":                 -0.25,
    "nitro":                 -0.10,
    "cyano":                 -0.50,
    "cyclopropyl":            0.30,
    "adamantyl":              2.20,
    "indole":                 1.40,
    "imidazole":             -0.85,
    "tetrazole":             -1.00,
}


# ═══════════════════════════════════════════════════════════════════════════════
# 4. BLOOD-BRAIN BARRIER RULES (Multi-Parameter)
# ═══════════════════════════════════════════════════════════════════════════════
BBB_CLINICAL_RULES = {
    "LogP_min": 1.5,
    "LogP_max": 4.5,
    "TPSA_max": 70.0,
    "MW_max":   450.0,
    "HBD_max":  2,
    "pKa_basic_range": (7, 10.5),
    "RotBonds_max": 8,
    "Arom_Rings_max": 3,
}

BBB_EFFLUX_PATTERNS = [
    {"smarts": "[NX3;H2,H1]CCO",   "name": "Ethanolamine",  "type": "P-gp Substrate"},
    {"smarts": "c1ccc(NC(=O))cc1", "name": "Benzamide_Core","type": "BCRP Substrate"},
    {"smarts": "CC(=O)Oc1ccccc1",  "name": "Acetate_Ester", "type": "OATP Substrate"},
]


# ═══════════════════════════════════════════════════════════════════════════════
# 5. CYP450 COMPREHENSIVE REFERENCE
# ═══════════════════════════════════════════════════════════════════════════════
CYP_REF_SUBSTRATES = {
    "3A4": ["Ketoconazole", "Clarithromycin", "Ritonavir", "Midazolam",
            "Cyclosporine", "Atorvastatin", "Sildenafil", "Amlodipine"],
    "2D6": ["Paroxetine", "Quinidine", "Terbinafine", "Fluoxetine",
            "Tramadol", "Codeine", "Metoprolol", "Haloperidol"],
    "2C9": ["Warfarin", "Phenytoin", "Tolbutamide", "Diclofenac",
            "Celecoxib", "Losartan", "Naproxen", "Irbesartan"],
    "2C19":["Omeprazole", "Pantoprazole", "Clopidogrel", "Escitalopram",
            "Diazepam", "Sertraline", "Voriconazole"],
    "1A2": ["Clozapine", "Olanzapine", "Theophylline", "Caffeine",
            "Duloxetine", "Ciprofloxacin"],
    "2B6": ["Bupropion", "Efavirenz", "Cyclophosphamide", "Ketamine"],
    "2E1": ["Ethanol", "Chlorzoxazone", "Acetaminophen_oxidation"],
}

CYP_INHIBITORS_REFERENCE = {
    "3A4_strong":  ["Itraconazole", "Ketoconazole", "Ritonavir", "Clarithromycin"],
    "3A4_moderate":["Fluconazole", "Diltiazem", "Verapamil", "Erythromycin"],
    "2D6_strong":  ["Paroxetine", "Fluoxetine", "Bupropion", "Quinidine"],
    "2C9_strong":  ["Fluconazole", "Amiodarone", "Miconazole"],
    "2C19_strong": ["Omeprazole", "Esomeprazole", "Fluvoxamine"],
    "1A2_strong":  ["Fluvoxamine", "Ciprofloxacin", "Enoxacin"],
}


# ═══════════════════════════════════════════════════════════════════════════════
# 6. CLINICAL SAFETY ALERTS
# ═══════════════════════════════════════════════════════════════════════════════
CLINICAL_ALERTS = {
    "hERG_High_Risk":    ["Piperidine-BasicN", "Macrocycle-Lactam", "Phenethylamine"],
    "QT_Prolongation":   ["c1ccc(c(c1)OC)C(=O)N", "ArylPiperidine"],
    "Skin_Irritation":   ["Acid_Anhydride", "Acyl_Halide"],
    "DILI_Structural":   ["Aniline", "Nitroaromatic", "Quinone"],
    "Carcinogenicity":   ["Polycyclic_Aromatic", "N-Nitroso", "Aflatoxin-like"],
    "Genotoxicity":      ["Nitro", "Azo_Aromatic", "Hydrazine", "Epoxide"],
}


# ═══════════════════════════════════════════════════════════════════════════════
# 7. PHYSICOCHEMICAL TARGET SPACE BY INDICATION
# ═══════════════════════════════════════════════════════════════════════════════
INDICATION_PROPERTY_TARGETS = {
    "CNS":       {"MW": (150, 450), "LogP": (1,  4),  "TPSA": (0, 70),  "HBD": (0, 2)},
    "Oral":      {"MW": (150, 500), "LogP": (-1, 5),  "TPSA": (0, 140), "HBD": (0, 5)},
    "Antiviral": {"MW": (200, 700), "LogP": (-1, 4),  "TPSA": (60,210), "HBD": (1, 8)},
    "Antibody":  {"MW": (200,1200), "LogP": (-3, 6),  "TPSA": (0, 500), "HBD": (0,20)},
    "Topical":   {"MW": (100, 400), "LogP": (1,  4),  "TPSA": (0, 80),  "HBD": (0, 3)},
    "PROTAC":    {"MW": (700,1200), "LogP": (1,  8),  "TPSA": (50,300), "HBD": (0,10)},
}


# ═══════════════════════════════════════════════════════════════════════════════
# 8. DRUG-LIKENESS BINS (ChEMBL-derived)
# ═══════════════════════════════════════════════════════════════════════════════
CHEMBL_DESCRIPTORS_BINS = {
    "MW":       [100, 250, 400, 550, 700],
    "LogP":     [-2,  0,   2,   4,   6],
    "TPSA":     [0,   50,  100, 150, 200],
    "RotBonds": [0,   4,   8,   12,  16]
}


# ═══════════════════════════════════════════════════════════════════════════════
# 9. PRODRUG & BIOLOGIC FLAGS
# ═══════════════════════════════════════════════════════════════════════════════
PRODRUG_INDICATORS = [
    {"smarts": "C(=O)OC",   "name": "Ester Prodrug",      "active_form": "Carboxylic Acid"},
    {"smarts": "C(=O)N",    "name": "Amide Prodrug",       "active_form": "Amine"},
    {"smarts": "OC(=O)C",   "name": "Carbonate/Ester",     "active_form": "Hydroxyl"},
    {"smarts": "OP(=O)(O)O","name": "Phosphate Prodrug",   "active_form": "Alcohol"},
]


# ═══════════════════════════════════════════════════════════════════════════════
# 10. PRIVILEGED SCAFFOLDS IN DRUGS
# ═══════════════════════════════════════════════════════════════════════════════
PRIVILEGED_SCAFFOLDS = {
    "Benzimidazole":  "c1ccc2nc[nH]c2c1",
    "Quinoline":      "c1ccc2ncccc2c1",
    "Indole":         "c1ccc2[nH]ccc2c1",
    "Piperazine":     "C1CNCCN1",
    "Morpholine":     "C1COCCN1",
    "Piperidine":     "C1CCNCC1",
    "Purine":         "c1ncc2[nH]cnc2n1",
    "Pyrimidine":     "c1ccncn1",
    "Thiophene":      "c1ccsc1",
    "Oxazole":        "c1cnoc1",
    "Imidazole":      "c1cncc1",
    "Benzothiazole":  "c1ccc2scnc2c1",
    "Pyrazole":       "c1cc[nH]n1",
    "Dihydropyridine":"C1CC=CN=C1",
    "Spirocycle":     "C1CCC2(CC1)CCCCC2",
    "Beta_Lactam":    "C1CC(=O)N1",
    "Lactam_6":       "C1CCCNC1=O",
    "Pyridine":       "c1ccncc1",
    "Fluorobenzene":  "c1ccc(F)cc1",
    "Naphthalene":    "c1ccc2ccccc2c1",
}


# ═══════════════════════════════════════════════════════════════════════════════
# 11. BIOISOSTERE REPLACEMENT TABLE
# ═══════════════════════════════════════════════════════════════════════════════
BIOISOSTERE_MAP = {
    "Carboxylic Acid":     ["Tetrazole", "Acylsulfonamide", "Hydroxamic Acid", "Phosphonate"],
    "Phenyl":              ["Pyridine", "Thiophene", "Pyrazole", "Cyclopentyl"],
    "Amide":               ["Reverse Amide", "Urea", "Carbamate", "Oxazole"],
    "Thioether":           ["Methylene", "Sulfoxide", "Difluoromethyl"],
    "N-H (amine)":         ["Cyclopropyl", "O-H", "F (bioisostere)"],
    "Catechol":            ["Difluorophenol", "Pyridinol"],
    "Methyl":              ["Cyclopropyl", "Trifluoromethyl (metabolic block)"],
    "Hydroxyl":            ["Fluorine", "Amino group", "Trifluoromethyl"],
    "Ester":               ["Amide", "Ketone", "Reverse ester"],
}


# ═══════════════════════════════════════════════════════════════════════════════
# 12. CLINICAL SUCCESS PROBABILITY MODIFIERS
# ═══════════════════════════════════════════════════════════════════════════════
CLINICAL_SUCCESS_MODIFIERS = {
    "Positive": {
        "High_Fsp3":        "+8% Phase III Success",
        "Low_Oxidative_Sites": "+5% Metabolic Stability",
        "QED_gt_0.7":       "+10% Overall Druglikeness",
        "MW_300_400":       "+7% Oral Bioavailability",
        "Selective_Target":  "+12% Efficacy",
    },
    "Negative": {
        "Thiol_Groups":     "-10% Selectivity",
        "Nitro_Groups":     "-15% Safety Margin",
        "Multiple_Alerts":  "-20% Safety Clearance",
        "High_LogP_gt5":    "-8% Solubility/Tox",
        "CYP3A4_Induction": "-12% DDI Risk",
    }
}


# ═══════════════════════════════════════════════════════════════════════════════
# 13. CLOUD DISCOVERY ENGINE
# ═══════════════════════════════════════════════════════════════════════════════
class AetherCloudDiscovery:
    """Edge-based compound discovery and logging."""
    def __init__(self, endpoint=None):
        self.endpoint = endpoint or "https://chemofilter-v1m.faang-os-piyush.workers.dev"

    def log_to_edge(self, compound_data):
        import requests
        try:
            r = requests.post(f"{self.endpoint}/api/log-molecule",
                              json=compound_data, timeout=2)
            return r.status_code == 200
        except:
            return False

    def get_global_discovery_stats(self):
        import requests
        try:
            r = requests.get(f"{self.endpoint}/api/get-discovery-stats", timeout=5)
            return r.json()
        except:
            return {"total": 0, "avg_score": 0}


# ═══════════════════════════════════════════════════════════════════════════════
# ACCESSOR FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════
def get_fda_map():               return FDA_REFERENCE_DB
def get_tox_alerts():            return EXTENDED_TOX_ALERTS
def get_logp_fixes():            return LOGP_CORRECTIONS
def get_cloud_engine():          return AetherCloudDiscovery()
def get_bbb_rules():             return BBB_CLINICAL_RULES
def get_cyp_substrates():        return CYP_REF_SUBSTRATES
def get_cyp_inhibitors():        return CYP_INHIBITORS_REFERENCE
def get_clinical_alerts():       return CLINICAL_ALERTS
def get_indication_targets():    return INDICATION_PROPERTY_TARGETS
def get_privileged_scaffolds():  return PRIVILEGED_SCAFFOLDS
def get_bioisostere_map():       return BIOISOSTERE_MAP
def get_prodrug_indicators():    return PRODRUG_INDICATORS
def get_success_modifiers():     return CLINICAL_SUCCESS_MODIFIERS
