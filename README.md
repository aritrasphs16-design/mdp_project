# ⬡ ChemoFilter — Crystalline Noir Edition
### VIT Chennai MDP 2026 · Omnipotent v1,000,000

A production-grade computational drug screening dashboard — amber gold on deep midnight navy, 21 active features, AI-powered analysis.

---

## 🚀 Deploy on Streamlit Cloud (share.streamlit.io)

### Step 1 — Push to GitHub
```bash
git init
git add .
git commit -m "ChemoFilter v1M"
git remote add origin https://github.com/YOUR_USERNAME/chemofilter.git
git push -u origin main
```

### Step 2 — Connect on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
2. Click **New app**
3. Select your repo, branch (`main`), and set **Main file path** to `app.py`
4. Click **Deploy**

### Step 3 — Add Your API Key (for AI features)
1. In the Streamlit Cloud dashboard, open your app → **⋮ menu** → **Settings**
2. Go to the **Secrets** tab
3. Paste this:
```toml
ANTHROPIC_API_KEY = "sk-ant-your-key-here"
```
4. Click **Save** — the app will reboot with AI features enabled

> Without the API key, the app still works fully — AI Explainer, Analogues, and Repurposing will just show a fallback message.

---

## 💻 Run Locally

```bash
# Install dependencies
pip install -r requirements.txt

# Add your API key for local dev
cp .streamlit/secrets.toml.template .streamlit/secrets.toml
# Edit secrets.toml and fill in ANTHROPIC_API_KEY

# Run
streamlit run app.py
```

---

## 📦 File Structure
```
chemofilter/
├── app.py                       ← Main app (Crystalline Noir, v1M)
├── requirements.txt             ← Python deps
├── .gitignore                   ← Keeps secrets.toml out of git
├── .streamlit/
│   ├── config.toml              ← Theme + server settings
│   └── secrets.toml.template   ← Copy → secrets.toml for local dev
│
├── features_v15.py              ← ADME features
├── mega_features_v20.py         ← 50+ extended features
├── quantum_accuracy_engine.py   ← LogP refinement, FDA similarity
├── hyper_zenith_v50.py          ← Hyper-zenith research module
├── master_drug_atlas.py         ← FDA drug atlas
├── chemical_intelligence_db.py  ← Core DB: toxicophores, CYP, BBB
├── omnipotent_engine_v200.py    ← Singularity engine
├── omnipotent_reactivity_db.py  ← Metabolic transformation DB
├── universal_analysis_v500.py   ← Organ tox, pharmacophore mapping
├── universal_blueprint_v500.py  ← Blueprint data
├── celestial_engine_v1000.py    ← Celestial engine
├── celestial_data_v1000.py
├── omega_engine_v2000.py        ← Omega-zenith engine
├── omega_data_v2000.py
├── xenon_engine_v5000.py        ← Xenon-god engine
├── xenon_data_v5000.py
├── aether_engine_v10000.py      ← Aether-primality engine (god-mode)
└── aether_data_v10000.py
```

---

## 🔑 Secrets Reference

| Key | Where to set | Purpose |
|-----|-------------|---------|
| `ANTHROPIC_API_KEY` | Streamlit Cloud → App Settings → Secrets | AI Explainer, Analogues, Repurposing |

---

## 📚 References
| # | Reference | Year |
|---|---|---|
|[1]|Daina & Zoete, ChemMedChem 11:1117|2016|
|[2]|Lipinski et al., ADDR 46:3|2001|
|[3]|Delaney, JCICS 44:1000|2004|
|[4]|Bickerton et al., Nat Chem 4:90|2012|
|[5]|Wager et al., ACS Chem Neurosci 1:435|2010|
|[6]|Baell & Holloway, JMC 53:2719|2010|
|[7]|Ertl & Schuffenhauer, J Cheminf 1:8|2009|
|[8]|Rogers & Hahn, JCIM 50:742|2010|
|[9]|Landrum, RDKit|2006+|
