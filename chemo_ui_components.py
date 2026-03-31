
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import pandas as pd
import numpy as np

def render_chemo_test_results(test_list):
    """
    Renders the results from the Vanguard Engine.
    test_list: list of dicts {"category": ..., "test": ..., "result": ..., "detail": ...}
    """
    if not test_list:
        st.info("No detailed test data available.")
        return

    # Category grouping
    cats = {}
    for t in test_list:
        c = t.get("category", "General")
        if c not in cats: cats[c] = []
        cats[c].append(t)
    
    # Define Tab groupings
    major_groups = {
        "🏗 Integrity": ["Structure Integrity"],
        "📏 PhysChem": ["Physicochemical"],
        "🏆 Guardrails": ["Drug-Likeness Rules", "Safety Catalogs"],
        "🧪 Structural Alerts": [] 
    }
    
    # Identify alert categories
    for c in cats.keys():
        if c.startswith("Alert:"):
            major_groups["🧪 Structural Alerts"].append(c)
    
    # Leftovers
    all_assigned = []
    for g in major_groups.values(): all_assigned.extend(g)
    other_cats = [c for c in cats.keys() if c not in all_assigned]
    if other_cats:
        major_groups["🔍 Deep Metrics"] = other_cats

    tabs = st.tabs(list(major_groups.keys()))
    
    for i, (tab_name, cat_list) in enumerate(major_groups.items()):
        with tabs[i]:
            if not cat_list:
                st.info(f"No {tab_name} data detected.")
                continue
            
            for cat in cat_list:
                if cat in cats:
                    st.markdown(f'<p style="font-size:0.75rem; color:var(--gold); letter-spacing:2px; margin:20px 0 10px 0">{cat.upper()}</p>', unsafe_allow_html=True)
                    cols = st.columns(3)
                    for j, test in enumerate(cats[cat]):
                        c = cols[j % 3]
                        res_str = test["result"]
                        detail = test["detail"]
                        name = test["test"]
                        
                        color = "#34d399" if res_str == "PASS" else "#f87171" if res_str == "FAIL" else "#fbbf24" if res_str == "WARN" else "#38bdf8"
                        icon = "✅" if res_str == "PASS" else "❌" if res_str == "FAIL" else "⚠️" if res_str == "WARN" else "ℹ️"
                        
                        c.markdown(f"""
                        <div style='padding:10px; background:rgba(255,255,255,0.02); border-radius:10px; margin-bottom:10px; border-left:3px solid {color}'>
                            <div style='font-size:0.65rem; color:var(--muted); text-transform:uppercase'>{name}</div>
                            <div style='font-size:0.85rem; color:white; font-weight:700'>{icon} {detail}</div>
                        </div>
                        """, unsafe_allow_html=True)

def render_filtering_explainer(test_list):
    """Parses the test list for critical FAIL or WARN results."""
    if not test_list: return
    criticals = [t for t in test_list if t["result"] in ["FAIL", "WARN"]]
    
    if criticals:
        st.markdown('<div class="rh">Critical Discovery Intelligence Alerts</div>', unsafe_allow_html=True)
        for c in criticals:
            if c["result"] == "FAIL":
                st.error(f"**{c['test']}**: {c['detail']} ({c['category']})")
            else:
                st.warning(f"**{c['test']}**: {c['detail']} ({c['category']})")
    else:
        st.success("✨ This molecule maintains absolute structural and med-chem integrity across all 800+ Vanguard validation tiers.")

def render_chemo_score_card(pkg):
    score = pkg["score"]
    grade = pkg["grade"]
    
    st.markdown(f"""
    <div style="background:linear-gradient(135deg, #0a1120, #142142); padding:40px; border-radius:24px; border:1px solid rgba(232,160,32,0.15); text-align:center; box-shadow:0 20px 50px rgba(0,0,0,0.5)">
        <div style="font-family:'DM Mono'; font-size:0.7rem; color:rgba(232,160,32,0.5); letter-spacing:6px; margin-bottom:10px">OMNIPOTENT LEAD SCORE V2.0</div>
        <div style="font-family:'Instrument Serif'; font-size:6.5rem; font-weight:400; color:white; line-height:1">{score}</div>
        <div style="font-family:'DM Mono'; font-size:1.4rem; color:#f5a623; letter-spacing:4px; margin-top:10px">GRADE: {grade}</div>
    </div>
    """, unsafe_allow_html=True)
    
    comp = pkg["components"]
    fig = go.Figure(go.Bar(
        x=list(comp.keys()),
        y=list(comp.values()),
        marker=dict(color=list(comp.values()), colorscale="YlOrBr", line=dict(color="rgba(255,255,255,0.1)", width=1))
    ))
    fig.update_layout(
        title=dict(text="Intelligence Weight Distribution", font=dict(size=14, color="rgba(245,166,35,0.6)")),
        template="plotly_dark",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font_color='rgba(200,222,255,0.4)',
        yaxis_range=[0, 1.1],
        height=300,
        margin=dict(l=40,r=40,t=60,b=40)
    )
    st.plotly_chart(fig, use_container_width=True)

def render_batch_intelligence(intel):
    """Renders the comprehensive dataset intelligence dashboard."""
    if not intel: 
        st.info("No dataset intelligence data available. Run analysis to populate.")
        return
    
    st.markdown('<div class="rh">Global Discovery Intelligence Hub</div>', unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Compound Count", intel.get("count", 0))
    c2.metric("Diversity Score", f"{intel.get('diversity_score', 0):.3f}")
    c3.metric("Avg MW", intel.get("mw_avg", 0))
    c4.metric("Avg LogP", intel.get("logp_avg", 0))
    
    st.markdown("---")
    
    t_sum, t_rules, t_scaf = st.tabs(["📊 Population Summary", "🏆 Rule Distribution", "🏗 Scaffold Analysis"])
    
    with t_sum:
        st.markdown("**Core Physicochemical Statistics**")
        ranges = intel.get("property_ranges", {})
        if ranges:
            st.table(pd.DataFrame([ranges], index=["Range Information"]))
        
    with t_rules:
        st.markdown("**Dataset Grade Distribution**")
        grades = intel.get("grade_dist", {})
        if grades:
            fig = px.pie(names=list(grades.keys()), values=list(grades.values()), 
                         template="plotly_dark", hole=0.4, color_discrete_sequence=px.colors.sequential.YlOrBr)
            fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No grade distribution data.")
            
    with t_scaf:
        st.markdown("**Top Molecular Scaffolds (Murcko)**")
        scafs = intel.get("top_scaffolds", [])
        if scafs:
            st.dataframe(pd.DataFrame(scafs), use_container_width=True, hide_index=True)
        else:
            st.info("No scaffolds detected.")
            
    st.markdown("**Functional Motif & Structural Alert Landscape**")
    fg_dist = intel.get("fg_distribution", {})
    if fg_dist:
        fg_df = pd.DataFrame({"Group": list(fg_dist.keys()), "Count": list(fg_dist.values())}).sort_values("Count", ascending=False).head(20)
        fig = px.bar(fg_df, x="Count", y="Group", orientation='h', color="Count",
                     color_continuous_scale="YlOrBr", template="plotly_dark")
        fig.update_layout(height=500, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                          font_color='rgba(200,222,255,0.4)', margin=dict(l=0,r=0,t=0,b=0))
        st.plotly_chart(fig, use_container_width=True)

def plot_chemo_property_distribution(df, prop, title):
    if df.empty or prop not in df: return None
    import plotly.express as px
    fig = px.histogram(df, x=prop, title=title, 
                       color_discrete_sequence=['#f5a623'],
                       template="plotly_dark",
                       nbins=40)
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font_color='rgba(200,222,255,0.4)',
        bargap=0.1
    )
    return fig

def plot_chemo_scatter(df, x_prop, y_prop):
    if df.empty or x_prop not in df or y_prop not in df: return None
    import plotly.express as px
    fig = px.scatter(df, x=x_prop, y=y_prop, color="Grade", 
                     title=f"{x_prop} vs {y_prop} Insight",
                     template="plotly_dark",
                     color_discrete_sequence=px.colors.qualitative.Bold)
    fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
    return fig


def render_preset_selector():
    """Renders a preset mode selector in the sidebar and returns the selected mode name."""
    import streamlit as st
    modes = ["Drug-Like (Ro5)", "Fragment-Like", "CNS Penetrant", "Natural Product-Like", "Custom"]
    mode = st.selectbox("Filter Preset", modes, index=0, key="preset_mode_selector")
    return mode

def get_preset_parameters(mode):
    """Returns slider default parameters dict for the given preset mode."""
    presets = {
        "Drug-Like (Ro5)":       {"mw_max": 500,  "logp_min": -2.0, "logp_max": 5.0,  "tpsa_max": 140, "hbd_max": 5,  "hba_max": 10, "rot_max": 10},
        "Fragment-Like":          {"mw_max": 300,  "logp_min": -2.0, "logp_max": 3.0,  "tpsa_max": 100, "hbd_max": 3,  "hba_max": 6,  "rot_max": 5},
        "CNS Penetrant":          {"mw_max": 450,  "logp_min": 0.0,  "logp_max": 5.0,  "tpsa_max": 90,  "hbd_max": 3,  "hba_max": 7,  "rot_max": 8},
        "Natural Product-Like":   {"mw_max": 600,  "logp_min": -3.0, "logp_max": 6.0,  "tpsa_max": 160, "hbd_max": 6,  "hba_max": 12, "rot_max": 12},
        "Custom":                 {"mw_max": 800,  "logp_min": -5.0, "logp_max": 10.0, "tpsa_max": 250, "hbd_max": 15, "hba_max": 20, "rot_max": 20},
    }
    return presets.get(mode, presets["Drug-Like (Ro5)"])
