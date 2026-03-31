
import pandas as pd
import json

def export_to_csv(data_list):
    df = pd.DataFrame(data_list)
    # Remove large internal objects before export
    cols_to_keep = [c for c in df.columns if not c.startswith("_")]
    return df[cols_to_keep].to_csv(index=False).encode('utf-8')

def export_to_json(data_list):
    df = pd.DataFrame(data_list)
    cols_to_keep = [c for c in df.columns if not c.startswith("_")]
    return df[cols_to_keep].to_json(orient="records", indent=4).encode('utf-8')

def generate_text_report(data_list):
    if not data_list: return "No data."
    
    df = pd.DataFrame(data_list)
    report = "CHEMOFILTER BATCH REPORT\n"
    report += "========================\n\n"
    report += f"Total Molecules: {len(df)}\n"
    report += f"Average ChemoScore: {df['ChemoScore'].mean():.2f}\n\n"
    
    report += "TOP SCORING LEADS\n"
    report += "-----------------\n"
    top = df.sort_values("ChemoScore", ascending=False).head(10)
    for _, row in top.iterrows():
        report += f"{row['ID']} ({row['Grade']}): Score {row['ChemoScore']} | SMILES: {row['SMILES']}\n"
            
    return report.encode('utf-8')
