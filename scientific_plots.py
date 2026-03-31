"""
╔══════════════════════════════════════════════════════════════════════════════╗
║   CHEMOFILTER — SCIENTIFIC PLOTS & EXPERIMENTAL WORKFLOW                    ║
║   Plotly Charts · Energy Profiles · Rate Plots · Experiment Builder         ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import math
import hashlib


def _seed(smiles):
    return int(hashlib.md5(smiles.encode()).hexdigest(), 16) % 10**8


# ═══════════════════════════════════════════════════════════════════
# SCIENTIFIC PLOT GENERATORS (Plotly)
# ═══════════════════════════════════════════════════════════════════

PLOT_THEME = dict(
    paper_bgcolor="rgba(5,8,15,0)",
    plot_bgcolor="rgba(5,8,15,0.4)",
    font=dict(family="IBM Plex Mono, monospace", size=12, color="#c8deff"),
    title_font=dict(family="DM Serif Display, serif", size=18, color="#e8a020"),
    legend=dict(bgcolor="rgba(5,8,15,0.6)", bordercolor="rgba(232,160,32,0.15)", borderwidth=1),
    xaxis=dict(gridcolor="rgba(200,222,255,0.05)", zerolinecolor="rgba(200,222,255,0.08)"),
    yaxis=dict(gridcolor="rgba(200,222,255,0.05)", zerolinecolor="rgba(200,222,255,0.08)"),
)


def apply_theme(fig):
    fig.update_layout(**PLOT_THEME)
    fig.update_layout(margin=dict(l=40, r=20, t=50, b=40))
    return fig


# ── 1. Reaction Energy Profile ──────────────────────────────────

def plot_reaction_energy_profile(smiles, ea=None, delta_h=None):
    """Reaction coordinate diagram with activation energy and enthalpy."""
    seed = _seed(smiles)
    if ea is None:
        ea = 40 + (seed % 60)
    if delta_h is None:
        delta_h = -30 + (seed % 60)

    x = np.linspace(0, 10, 100)
    # Gaussian barrier + product shift
    profile = np.zeros_like(x)
    profile += ea * np.exp(-0.5 * ((x - 4) / 1.2) ** 2)  # Barrier
    profile += delta_h * (1 / (1 + np.exp(-2 * (x - 6))))  # Product shift

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=profile, mode="lines",
        line=dict(color="#e8a020", width=3),
        fill="tozeroy", fillcolor="rgba(232,160,32,0.05)",
        name="Energy Profile"
    ))

    # Annotations
    fig.add_annotation(x=4, y=ea, text=f"Ea = {ea} kJ/mol",
                       showarrow=True, arrowcolor="#f87171", font=dict(color="#f87171"))
    fig.add_annotation(x=8, y=delta_h, text=f"ΔH = {delta_h} kJ/mol",
                       showarrow=True, arrowcolor="#4ade80", font=dict(color="#4ade80"))

    fig.update_layout(title="Reaction Energy Profile",
                      xaxis_title="Reaction Coordinate",
                      yaxis_title="Energy (kJ/mol)")
    return apply_theme(fig)


# ── 2. Rate vs Temperature ──────────────────────────────────────

def plot_rate_vs_temperature(smiles, ea=None):
    """Arrhenius-style rate constant vs temperature plot."""
    seed = _seed(smiles)
    if ea is None:
        ea = 40 + (seed % 60)
    A = 1e10

    temps = np.linspace(250, 450, 50)
    R = 8.314
    rates = A * np.exp(-ea * 1000 / (R * temps))

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=temps - 273.15, y=np.log10(rates), mode="lines+markers",
        line=dict(color="#38bdf8", width=2),
        marker=dict(size=4, color="#38bdf8"),
        name="ln(k)"
    ))

    fig.update_layout(title=f"Rate vs Temperature (Ea = {ea} kJ/mol)",
                      xaxis_title="Temperature (°C)",
                      yaxis_title="log₁₀(Rate Constant)")
    return apply_theme(fig)


# ── 3. Concentration vs Time ────────────────────────────────────

def plot_concentration_vs_time(smiles, k=None, c0=1.0):
    """First-order decay concentration vs time."""
    seed = _seed(smiles)
    if k is None:
        k = 0.01 + (seed % 100) / 1000

    t = np.linspace(0, 500, 100)
    c = c0 * np.exp(-k * t)
    product = c0 - c

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=t, y=c, mode="lines", name="Reactant",
                             line=dict(color="#f87171", width=2)))
    fig.add_trace(go.Scatter(x=t, y=product, mode="lines", name="Product",
                             line=dict(color="#4ade80", width=2)))

    # Half-life line
    t_half = 0.693 / k
    fig.add_vline(x=t_half, line_dash="dot", line_color="#fbbf24",
                  annotation_text=f"t½ = {t_half:.1f}s")

    fig.update_layout(title=f"Concentration vs Time (k = {k:.4f} s⁻¹)",
                      xaxis_title="Time (s)",
                      yaxis_title="Concentration (M)")
    return apply_theme(fig)


# ── 4. Yield vs Catalyst Comparison ─────────────────────────────

def plot_yield_vs_catalyst(catalyst_data):
    """Bar chart comparing yields across catalysts."""
    names = [c["Catalyst"] for c in catalyst_data[:8]]
    yields = [float(c["Est_Yield"].replace("%", "")) for c in catalyst_data[:8]]
    colors = ["#4ade80" if y > 80 else "#fbbf24" if y > 60 else "#f87171" for y in yields]

    fig = go.Figure([go.Bar(x=names, y=yields, marker_color=colors,
                            text=[f"{y}%" for y in yields], textposition="auto")])

    fig.update_layout(title="Yield Comparison by Catalyst",
                      xaxis_title="Catalyst", yaxis_title="Yield (%)",
                      yaxis=dict(range=[0, 105]))
    return apply_theme(fig)


# ── 5. pH Stability Surface ─────────────────────────────────────

def plot_ph_stability(ph_data):
    """Line chart of stability vs pH."""
    phs = [d["pH"] for d in ph_data]
    stab = [d["Stability_%"] for d in ph_data]

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=phs, y=stab, mode="lines+markers",
                             line=dict(color="#a78bfa", width=2, shape="spline"),
                             marker=dict(size=8, color="#a78bfa"),
                             fill="tozeroy", fillcolor="rgba(167,139,250,0.05)",
                             name="Stability"))

    # Optimal zone
    fig.add_vrect(x0=5, x1=8, fillcolor="rgba(52,211,153,0.05)",
                  line_width=0, annotation_text="Optimal Zone")

    fig.update_layout(title="Stability vs pH", xaxis_title="pH",
                      yaxis_title="Stability (%)", yaxis=dict(range=[0, 105]))
    return apply_theme(fig)


# ── 6. Solvent Comparison Radar ──────────────────────────────────

def plot_solvent_radar(solvent_data):
    """Radar chart comparing top solvents."""
    top5 = solvent_data[:5]
    categories = ["Polarity", "BP Score", "Yield"]

    fig = go.Figure()
    for s in top5:
        fig.add_trace(go.Scatterpolar(
            r=[s["Polarity"] / 10, s["Dielectric"] / 80,
               float(s["Est_Yield"].replace("%", "")) / 100],
            theta=categories, fill="toself", name=s["Solvent"],
            line=dict(width=2),
        ))

    fig.update_layout(
        title="Solvent Comparison",
        polar=dict(
            bgcolor="rgba(5,8,15,0.4)",
            angularaxis=dict(gridcolor="rgba(200,222,255,0.1)"),
            radialaxis=dict(gridcolor="rgba(200,222,255,0.08)", range=[0, 1]),
        ),
    )
    return apply_theme(fig)


# ── 7. Radar Chart for Compound Properties ──────────────────────

def plot_compound_radar(radar_data, compound_id="Cpd-1"):
    """5-axis radar chart for a single compound."""
    categories = list(radar_data.keys())
    values = list(radar_data.values())
    values.append(values[0])  # Close the polygon
    categories.append(categories[0])

    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=values, theta=categories, fill="toself",
        line=dict(color="#e8a020", width=2),
        fillcolor="rgba(232,160,32,0.12)",
        name=compound_id
    ))

    fig.update_layout(
        title=f"Property Radar — {compound_id}",
        polar=dict(
            bgcolor="rgba(5,8,15,0.4)",
            angularaxis=dict(gridcolor="rgba(200,222,255,0.1)"),
            radialaxis=dict(gridcolor="rgba(200,222,255,0.08)", range=[0, 1]),
        ),
    )
    return apply_theme(fig)


# ── 8. Bond Strength Distribution ───────────────────────────────

def plot_bond_strength_distribution(bond_data):
    """Histogram of bond dissociation energies."""
    bdes = [b["BDE_kJ_mol"] for b in bond_data["All_Bonds"]]

    fig = go.Figure()
    fig.add_trace(go.Histogram(
        x=bdes, nbinsx=15,
        marker_color="rgba(232,160,32,0.6)",
        marker_line=dict(color="#e8a020", width=1),
    ))
    fig.add_vline(x=bond_data["Average_BDE"], line_dash="dash",
                  line_color="#4ade80", annotation_text=f"Avg: {bond_data['Average_BDE']} kJ/mol")

    fig.update_layout(title="Bond Strength Distribution",
                      xaxis_title="BDE (kJ/mol)", yaxis_title="Count")
    return apply_theme(fig)


# ── 9. Gasteiger Charge Scatter ──────────────────────────────────

def plot_charge_scatter(charge_data):
    """Scatter plot of atomic charges."""
    atoms = charge_data["atoms"]
    if not atoms:
        return go.Figure()

    x = [a["idx"] for a in atoms]
    y = [a["charge"] for a in atoms]
    labels = [a["symbol"] for a in atoms]
    colors = ["#f87171" if c > 0.1 else "#38bdf8" if c < -0.1 else "#666" for c in y]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=y, mode="markers+text", text=labels,
        textposition="top center", textfont=dict(size=9),
        marker=dict(size=12, color=colors, line=dict(width=1, color="rgba(255,255,255,0.2)")),
    ))
    fig.add_hline(y=0, line_dash="dash", line_color="rgba(200,222,255,0.2)")

    fig.update_layout(title="Gasteiger Charge Distribution",
                      xaxis_title="Atom Index", yaxis_title="Partial Charge")
    return apply_theme(fig)


# ── 10. Kinetics Data Plot ──────────────────────────────────────

def plot_kinetics(kinetics_data):
    """Temperature-dependent kinetics plot."""
    data = kinetics_data.get("Data", [])
    if not data:
        return go.Figure()

    temps = [d["Temperature_C"] for d in data]
    rates = [float(d["Rate_Constant_k"]) for d in data]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=temps, y=rates, mode="lines+markers",
        line=dict(color="#34d399", width=2),
        marker=dict(size=6, color="#34d399"),
        name="Rate Constant"
    ))

    fig.update_layout(title=f"Kinetics — Ea = {kinetics_data['Activation_Energy_kJ']} kJ/mol",
                      xaxis_title="Temperature (°C)",
                      yaxis_title="Rate Constant (s⁻¹)",
                      yaxis_type="log")
    return apply_theme(fig)


# ── 11. Repeatability Scatter ───────────────────────────────────

def plot_repeatability(repeat_data):
    """Box + scatter of repeated runs."""
    runs = repeat_data.get("Individual_Runs", [])
    if not runs:
        return go.Figure()

    fig = go.Figure()
    fig.add_trace(go.Box(
        y=runs, name="Yield Distribution",
        marker_color="#e8a020",
        boxpoints="all", jitter=0.3, pointpos=-1.5,
    ))

    fig.add_hline(y=repeat_data["Mean_Yield"], line_dash="dash",
                  line_color="#4ade80", annotation_text=f"Mean: {repeat_data['Mean_Yield']}%")

    fig.update_layout(title=f"Repeatability ({repeat_data['Runs']} runs, CV={repeat_data['CV_Pct']}%)",
                      yaxis_title="Yield (%)")
    return apply_theme(fig)


# ── 12. Sensitivity Tornado ─────────────────────────────────────

def plot_sensitivity_tornado(sensitivity_data):
    """Tornado chart showing parameter sensitivity."""
    params = sensitivity_data.get("Parameters", [])
    if not params:
        return go.Figure()

    names = [p["Parameter"] for p in params]
    impacts = [p["Yield_Impact_%"] for p in params]
    colors = ["#f87171" if p["Classification"] == "Critical" else "#fbbf24" if p["Classification"] == "Sensitive" else "#4ade80" for p in params]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=names, x=impacts, orientation="h",
        marker_color=colors,
        text=[f"{v}%" for v in impacts], textposition="auto",
    ))

    fig.update_layout(title="Parameter Sensitivity (Tornado)",
                      xaxis_title="Yield Impact (%)")
    return apply_theme(fig)


# ── 13. Multi-Compound Comparison Heatmap ────────────────────────

def plot_compound_heatmap(compounds_data, properties=None):
    """Heatmap comparing multiple compounds on key properties."""
    if properties is None:
        properties = ["MW", "LogP", "tPSA", "QED", "LeadScore"]

    ids = [c.get("ID", f"Cpd-{i}") for i, c in enumerate(compounds_data)]
    z = []
    for prop in properties:
        row = []
        for c in compounds_data:
            val = c.get(prop, 0)
            try:
                val = float(val)
            except:
                val = 0
            row.append(val)
        z.append(row)

    fig = go.Figure(data=go.Heatmap(
        z=z, x=ids, y=properties,
        colorscale=[[0, "#1a1a2e"], [0.5, "#e8a020"], [1, "#4ade80"]],
        text=[[f"{v:.1f}" for v in row] for row in z],
        texttemplate="%{text}",
        textfont=dict(size=10),
    ))

    fig.update_layout(title="Compound Property Heatmap")
    return apply_theme(fig)


# ── 14. Drug-Likeness Spider (Multi) ────────────────────────────

def plot_multi_radar(compounds_data, max_compounds=4):
    """Overlay radar charts for multiple compounds."""
    colors = ["#e8a020", "#38bdf8", "#a78bfa", "#34d399", "#f87171"]

    fig = go.Figure()
    for i, c in enumerate(compounds_data[:max_compounds]):
        radar = c.get("_ext", {}).get("Radar_Data", {})
        if not radar:
            continue
        cats = list(radar.keys()) + [list(radar.keys())[0]]
        vals = list(radar.values()) + [list(radar.values())[0]]

        fig.add_trace(go.Scatterpolar(
            r=vals, theta=cats, fill="toself",
            name=c.get("ID", f"Cpd-{i+1}"),
            line=dict(color=colors[i % len(colors)], width=2),
        ))

    fig.update_layout(
        title="Multi-Compound Radar",
        polar=dict(
            bgcolor="rgba(5,8,15,0.4)",
            angularaxis=dict(gridcolor="rgba(200,222,255,0.1)"),
            radialaxis=dict(gridcolor="rgba(200,222,255,0.08)", range=[0, 1]),
        ),
    )
    return apply_theme(fig)


# ── 15. Equilibrium Diagram ─────────────────────────────────────

def plot_equilibrium(eq_data):
    """Bar chart showing products vs reactants at equilibrium."""
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=["Reactants", "Products"],
        y=[eq_data["Reactant_Pct"], eq_data["Product_Pct"]],
        marker_color=["#f87171", "#4ade80"],
        text=[f"{eq_data['Reactant_Pct']}%", f"{eq_data['Product_Pct']}%"],
        textposition="auto",
    ))

    fig.update_layout(
        title=f"Equilibrium Distribution (K = {eq_data['Equilibrium_K']})",
        yaxis_title="Percentage (%)", yaxis=dict(range=[0, 105]),
    )
    return apply_theme(fig)


# ═══════════════════════════════════════════════════════════════════
# EXPERIMENTAL WORKFLOW FEATURES
# ═══════════════════════════════════════════════════════════════════

def build_experiment(mol, params):
    """Build a complete experiment configuration."""
    from rdkit import Chem
    from rdkit.Chem import Descriptors
    smiles = Chem.MolToSmiles(mol) if hasattr(mol, 'GetNumAtoms') else str(mol)

    return {
        "Experiment_ID": f"EXP-{_seed(smiles) % 10000:04d}",
        "SMILES": smiles,
        "Temperature": params.get("temperature", 25),
        "Pressure": params.get("pressure", 1.0),
        "Solvent": params.get("solvent", "DMSO"),
        "Catalyst": params.get("catalyst", "None"),
        "Concentration": params.get("concentration", 0.1),
        "pH": params.get("pH", 7.0),
        "Duration": params.get("duration", "2h"),
        "Status": "Ready",
    }


def compare_scenarios(experiments):
    """Compare multiple experiment scenarios side by side."""
    comparison = []
    for exp in experiments:
        comparison.append({
            "ID": exp["Experiment_ID"],
            "Conditions": f"{exp['Temperature']}°C / {exp['Pressure']} atm / pH {exp['pH']}",
            "Solvent": exp["Solvent"],
            "Catalyst": exp["Catalyst"],
        })
    return comparison


def batch_testing_summary(results_list):
    """Summarize batch testing across multiple compounds."""
    if not results_list:
        return {}

    n = len(results_list)
    avg_yield = sum(float(r.get("Est_Yield", "0").replace("%", "")) for r in results_list) / n

    return {
        "Total_Compounds": n,
        "Average_Yield": round(avg_yield, 1),
        "Best_Performer": max(results_list, key=lambda r: float(r.get("Est_Yield", "0").replace("%", ""))),
        "Worst_Performer": min(results_list, key=lambda r: float(r.get("Est_Yield", "0").replace("%", ""))),
    }


def parameter_sweep_data(smiles, param_name="temperature", param_range=None):
    """Generate data for a parameter sweep visualization."""
    seed = _seed(smiles)
    if param_range is None:
        param_range = list(range(0, 101, 5))

    base = 60 + seed % 20
    optimal = param_range[len(param_range) // 2]

    data = []
    for val in param_range:
        yield_val = max(10, min(99, round(base - 0.02 * (val - optimal) ** 2 + (seed % 5), 1)))
        data.append({"Param_Value": val, "Yield_%": yield_val})

    return {
        "Parameter": param_name,
        "Range": f"{param_range[0]}-{param_range[-1]}",
        "Optimal_Value": optimal,
        "Data": data,
    }


def generate_report_text(compound_data, extended_data):
    """Generate automated text report for a compound."""
    lines = []
    lines.append(f"═══ ChemoFilter Automated Report ═══")
    lines.append(f"Compound: {compound_data.get('ID', 'Unknown')}")
    lines.append(f"SMILES: {compound_data.get('SMILES', 'N/A')}")
    lines.append(f"Grade: {compound_data.get('Grade', 'N/A')}")
    lines.append(f"Lead Score: {compound_data.get('LeadScore', 'N/A')}/100")
    lines.append("")
    lines.append("── Physicochemical Properties ──")
    lines.append(f"  MW: {compound_data.get('MW', 'N/A')} Da")
    lines.append(f"  LogP: {compound_data.get('LogP', 'N/A')}")
    lines.append(f"  tPSA: {compound_data.get('tPSA', 'N/A')} Å²")
    lines.append(f"  HBD/HBA: {extended_data.get('HBD', 'N/A')}/{extended_data.get('HBA', 'N/A')}")
    lines.append(f"  Rotatable Bonds: {extended_data.get('Rotatable_Bonds', 'N/A')}")
    lines.append(f"  Rings: {extended_data.get('Ring_Count', 'N/A')} (Aromatic: {extended_data.get('Aromatic_Ring_Count', 'N/A')})")
    lines.append("")
    lines.append("── ADMET Prediction ──")
    lines.append(f"  Solubility: {extended_data.get('Solubility_Class', 'N/A')} (LogS: {extended_data.get('LogS_ESOL', 'N/A')})")
    lines.append(f"  HIA: {extended_data.get('HIA', 'N/A')}")
    lines.append(f"  BBB: {extended_data.get('BBB_Penetration', 'N/A')}")
    lines.append(f"  CYP450 Risk: {extended_data.get('CYP450_Risk', 'N/A')}")
    lines.append(f"  Bioavailability: {extended_data.get('Bioavailability_Score', 'N/A')}")
    lines.append(f"  Toxicity Risk: {extended_data.get('Toxicity_Risk', 'N/A')}")
    lines.append(f"  Mutagenicity: {extended_data.get('Mutagenicity_Risk', 'N/A')}")
    lines.append("")
    lines.append("── Lead Optimization ──")
    lines.append(f"  Lead Potential: {extended_data.get('Lead_Optimization_Potential', 'N/A')}/100")
    lines.append(f"  Drug-Likeness: {extended_data.get('Drug_Likeness_Badge', 'N/A')}")
    lines.append(f"  Synthetic Difficulty: {extended_data.get('Synthetic_Difficulty', 'N/A')}")
    lines.append(f"  Optimization Priority: {extended_data.get('Optimization_Priority', 'N/A')}/100")

    if extended_data.get("Property_Warnings"):
        lines.append("")
        lines.append("── Warnings ──")
        for w in extended_data["Property_Warnings"]:
            lines.append(f"  {w}")

    return "\n".join(lines)


def plot_parameter_sweep(sweep_data):
    """Plot parameter sweep results."""
    data_pts = sweep_data.get("Data", [])
    if not data_pts:
        return go.Figure()

    x = [d["Param_Value"] for d in data_pts]
    y = [d["Yield_%"] for d in data_pts]

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=x, y=y, mode="lines+markers",
        line=dict(color="#e8a020", width=2, shape="spline"),
        marker=dict(size=6, color="#e8a020"),
        fill="tozeroy", fillcolor="rgba(232,160,32,0.05)",
    ))

    fig.add_vline(x=sweep_data["Optimal_Value"], line_dash="dot",
                  line_color="#4ade80",
                  annotation_text=f"Optimal: {sweep_data['Optimal_Value']}")

    fig.update_layout(title=f"Parameter Sweep — {sweep_data['Parameter']}",
                      xaxis_title=sweep_data["Parameter"],
                      yaxis_title="Yield (%)", yaxis=dict(range=[0, 105]))
    return apply_theme(fig)
