
import pandas as pd
import numpy as np
import time
import random
from rdkit import Chem
from rdkit.Chem import AllChem, DataStructs
from rdkit.ML.Cluster import Butina
from rdkit.Chem.Scaffolds import MurckoScaffold
import chemo_filters as cf
import chemo_scoring as cs

def process_molecule_batch(smis, ids=None):
    """Analyze a batch of SMILES and return a structured dataframe."""
    start = time.time()
    
    if not ids:
        ids = [f"Cpd-{i+1}" for i in range(len(smis))]
        
    results = []
    for smi, id_val in zip(smis, ids):
        try:
            mol = Chem.MolFromSmiles(smi)
            if mol:
                res = cf.run_comprehensive_screening(smi)
                res["ID"] = id_val
                res["SMILES"] = smi
                res["_mol"] = mol
                
                score_pkg = cs.get_chemoscore_pkg(res)
                res["ChemoScore"] = score_pkg["score"]
                res["Grade"] = score_pkg["grade"]
                res["_score_pkg"] = score_pkg
                
                results.append(res)
            else:
                results.append({"ID": id_val, "SMILES": smi, "valid": False, "ChemoScore": 0.0, "Grade": "F"})
        except Exception as e:
            print(f"Error processing {id_val}: {e}")
            continue
            
    df = pd.DataFrame(results)
    print(f"Batch processing took {time.time() - start:.2f}s")
    return df

def calculate_tanimoto_diversity(df):
    """Calculate the average pairwise Tanimoto similarity to measure diversity."""
    if df.empty or "SMILES" not in df: return 0.0
    
    mols = [Chem.MolFromSmiles(s) for s in df["SMILES"] if isinstance(s, str)]
    mols = [m for m in mols if m]
    if len(mols) < 2: return 0.0
    
    fps = [AllChem.GetMorganFingerprintAsBitVect(m, 2, nBits=1024) for m in mols]
    
    n = len(fps)
    sims = []
    if n > 150:
        sample_indices = random.sample(range(n), 150)
        fps = [fps[i] for i in sample_indices]
        n = 150
        
    for i in range(n):
        for j in range(i + 1, n):
            sims.append(DataStructs.TanimotoSimilarity(fps[i], fps[j]))
            
    avg_sim = np.mean(sims)
    return 1.0 - avg_sim

def get_scaffold_distribution(df):
    """Detect and count unique Murcko scaffolds in the batch."""
    if df.empty or "SMILES" not in df: return []
    
    scaffolds = []
    for s in df["SMILES"]:
        if not isinstance(s, str): continue
        try:
            m = Chem.MolFromSmiles(s)
            if m:
                scaff_mol = MurckoScaffold.GetScaffoldForMol(m)
                scaff_smi = Chem.MolToSmiles(scaff_mol)
                if scaff_smi:
                    scaffolds.append(scaff_smi)
        except:
            continue
            
    from collections import Counter
    counts = Counter(scaffolds)
    return [{"smiles": k, "count": v} for k, v in counts.most_common(10)]

def extract_dataset_intelligence(df):
    """Generate a high-fidelity dictionary of batch statistics for the Discovery Hub."""
    if df.empty: return {}
    
    # Standardize column naming just in case
    # If the df came from run_comprehensive_screening, it will have MW, LogP, etc.
    
    intel = {
        "count": len(df),
        "diversity_score": round(calculate_tanimoto_diversity(df), 3),
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        "status": "ANALYTICS-READY"
    }
    
    # Grade distribution
    if "Grade" in df:
        intel["grade_dist"] = df["Grade"].value_counts().to_dict()
        
    # Property stats
    for prop in ["MW", "LogP", "TPSA", "QED", "SA_Score"]:
        if prop in df:
            data = df[prop].dropna()
            if not data.empty:
                intel[f"{prop.lower()}_avg"] = round(float(data.mean()), 2)
                intel[f"{prop.lower()}_max"] = round(float(data.max()), 2)
                intel[f"{prop.lower()}_min"] = round(float(data.min()), 2)

    # Scaffolds
    intel["top_scaffolds"] = get_scaffold_distribution(df)
    
    # Structural Alert Hotspots
    fg_cols = [c for c in df.columns if c.startswith("has_")]
    if fg_cols:
        fg_sums = df[fg_cols].sum().sort_values(ascending=False)
        intel["fg_distribution"] = {k.replace("has_", "").replace("_", " ").title(): int(v) for k, v in fg_sums.items() if v > 0}
        
    # Lead Identification
    if "ChemoScore" in df:
        leads = df.sort_values("ChemoScore", ascending=False).head(10)
        # We need specific columns for the UI
        intel["leads"] = []
        for i, row in leads.iterrows():
            intel["leads"].append({
                "ID": row["ID"],
                "smiles": row.get("SMILES", ""),
                "Grade": row.get("Grade", "N/A"),
                "Score": round(float(row["ChemoScore"]), 2)
            })
            
    # Property Ranges for the UI table
    ranges = {}
    for prop in ["MW", "LogP", "TPSA", "HBD", "HBA", "RotBonds"]:
        if prop in df:
            data = df[prop].dropna()
            if not data.empty:
                ranges[prop] = f"{data.min():.1f} – {data.max():.1f}"
    intel["property_ranges"] = ranges

    return intel


def get_batch_statistics(data):
    """
    Compute summary statistics from a list of compound dicts (from analyze()).
    Returns a JSON-serialisable dict for st.json() display.
    """
    if not data:
        return {"count": 0, "message": "No data available"}

    import numpy as np

    stats = {"count": len(data)}
    numeric_keys = ["MW", "LogP", "TPSA", "QED", "SA_Score", "HBD", "HBA", "RotBonds", "Fsp3"]

    for key in numeric_keys:
        vals = []
        for d in data:
            v = d.get(key)
            try:
                vals.append(float(v))
            except (TypeError, ValueError):
                pass
        if vals:
            stats[f"{key}_avg"]  = round(float(np.mean(vals)), 3)
            stats[f"{key}_min"]  = round(float(np.min(vals)),  3)
            stats[f"{key}_max"]  = round(float(np.max(vals)),  3)
            stats[f"{key}_std"]  = round(float(np.std(vals)),  3)

    # Grade distribution
    grades = [d.get("Grade", "?") for d in data]
    from collections import Counter
    stats["grade_distribution"] = dict(Counter(grades))

    # Rule compliance
    for rule in ["Lipinski", "Veber", "Ghose", "Egan", "Muegge"]:
        compliant = sum(1 for d in data if d.get(rule) is True or
                        (isinstance(d.get("rules"), dict) and d["rules"].get(rule)))
        stats[f"{rule}_compliant"] = compliant
        stats[f"{rule}_pct"] = round(100 * compliant / len(data), 1)

    return stats
