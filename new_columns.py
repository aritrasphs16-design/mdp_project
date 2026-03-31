"""
new_columns.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · New Smart Column System
────────────────────────────────────────────────────────────────────────────

ADD-ONLY module. Injects 5 new smart columns into the results display.
Never modifies existing columns or logic.

COLUMNS ADDED
    🔬 Visualization   → link to visualization_app with SMILES
    📊 Data Portal     → link to data_portal with SMILES
    🧠 Insight         → one-line compound intelligence summary
    ⚠️  Risk            → key risk flags (hERG, PAINS, DILI…)
    🎯 Recommendation  → suggested next development step
────────────────────────────────────────────────────────────────────────────
"""
from __future__ import annotations
import urllib.parse

try:
    import streamlit as st
    _ST = True
except ImportError:
    _ST = False

try:
    import pandas as pd
    _PD = True
except ImportError:
    _PD = False

# ─────────────────────────────────────────────────────────────────────────────
# URL BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

_VIZ_BASE    = "http://localhost:8501/visualization"
_PORTAL_BASE = "http://localhost:8501/data_portal"

def get_visualization_url(smiles: str, base: str = _VIZ_BASE) -> str:
    """Build deep-link URL to visualization app for a given SMILES."""
    encoded = urllib.parse.quote(smiles.strip(), safe="")
    return f"{base}?smiles={encoded}"

def get_portal_url(smiles: str, base: str = _PORTAL_BASE) -> str:
    """Build deep-link URL to data portal for a given SMILES."""
    encoded = urllib.parse.quote(smiles.strip(), safe="")
    return f"{base}?smiles={encoded}"


# ─────────────────────────────────────────────────────────────────────────────
# INTELLIGENCE FUNCTIONS (pure, no side effects, always safe)
# ─────────────────────────────────────────────────────────────────────────────

def get_insight_summary(compound: dict) -> str:
    """
    Generate a one-line intelligence summary for a compound.
    Uses existing fields — no recomputation.
    """
    try:
        score = compound.get("LeadScore") or compound.get("lead_score", 0)
        qed   = compound.get("QED") or compound.get("qed", 0)
        grade = compound.get("Grade") or compound.get("grade", "?")
        mw    = compound.get("MW") or compound.get("mw", 0)
        bbb   = compound.get("_bbb") or compound.get("bbb_predicted", False)
        hia   = compound.get("_hia") or compound.get("hia_predicted", False)

        parts = [f"Grade {grade}"]
        if score:
            parts.append(f"LeadScore {float(score):.0f}")
        if qed:
            parts.append(f"QED {float(qed):.3f}")
        if bbb:
            parts.append("BBB+")
        if hia:
            parts.append("HIA+")
        if mw and float(mw) <= 300:
            parts.append("Fragment-like")
        elif mw and float(mw) <= 500:
            parts.append("Drug-like MW")

        return " · ".join(parts) if parts else "No insight available"
    except Exception:
        return "—"


def get_risk_summary(compound: dict) -> str:
    """
    Return a concise risk flag string.
    Uses existing fields — no recomputation.
    """
    try:
        risks = []
        herg  = str(compound.get("_herg", "") or compound.get("herg_risk", "")).upper()
        pains = int(compound.get("_pains", 0) or compound.get("pains_count", 0))
        dili  = compound.get("dili_risk", False)
        brenk = int(compound.get("brenk_count", 0))
        alerts= int(compound.get("total_alert_count", 0) or
                    compound.get("CYP_Hits", 0))
        logp  = float(compound.get("LogP") or compound.get("logp", 0))
        mw    = float(compound.get("MW") or compound.get("mw", 0))

        if "HIGH" in herg: risks.append("🔴 hERG")
        elif "MED" in herg or "MEDIUM" in herg: risks.append("🟡 hERG?")
        if pains > 0:  risks.append(f"🔴 PAINS×{pains}")
        if brenk > 0:  risks.append(f"🟡 Brenk×{brenk}")
        if dili:       risks.append("🔴 DILI")
        if logp > 5:   risks.append("🟡 LogP>5")
        if mw > 500:   risks.append("🟡 MW>500")
        if alerts > 3: risks.append(f"🔴 Alerts×{alerts}")

        return "  ".join(risks) if risks else "🟢 Clean"
    except Exception:
        return "—"


def get_recommendation(compound: dict) -> str:
    """
    Suggest next development step based on compound profile.
    Fast heuristic — no recomputation.
    """
    try:
        score = float(compound.get("LeadScore") or compound.get("lead_score", 0))
        pains = int(compound.get("_pains", 0) or compound.get("pains_count", 0))
        herg  = str(compound.get("_herg", "") or compound.get("herg_risk", "")).upper()
        grade = str(compound.get("Grade") or compound.get("grade", "D"))
        mw    = float(compound.get("MW") or compound.get("mw", 0))
        logp  = float(compound.get("LogP") or compound.get("logp", 0))
        dili  = compound.get("dili_risk", False)

        if grade == "A" and pains == 0:
            return "🚀 Progress to lead optimisation"
        if "HIGH" in herg:
            return "⚠️ Address hERG liability first"
        if pains > 1:
            return "🔧 Remove PAINS alerts"
        if dili:
            return "⚠️ Assess hepatotoxicity risk"
        if mw > 550:
            return "✂️ Reduce molecular weight"
        if logp > 5.5:
            return "💧 Improve aqueous solubility"
        if score >= 55:
            return "🔬 Explore structural analogues"
        if score >= 40:
            return "🧪 Scaffold hop or fragment merge"
        return "🔄 Redesign core scaffold"
    except Exception:
        return "—"


def get_development_priority(compound: dict) -> str:
    """Return HIGH / MEDIUM / LOW priority tag."""
    try:
        score = float(compound.get("LeadScore") or compound.get("lead_score", 0))
        pains = int(compound.get("_pains", 0) or compound.get("pains_count", 0))
        herg  = str(compound.get("_herg", "") or compound.get("herg_risk", "")).upper()
        if score >= 70 and pains == 0 and "HIGH" not in herg:
            return "⬆ HIGH"
        if score >= 50:
            return "➡ MEDIUM"
        return "⬇ LOW"
    except Exception:
        return "—"


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN INJECTION — called from app.py results section (ADD-ONLY)
# ─────────────────────────────────────────────────────────────────────────────

_NEW_COL_CSS = """<style>
.nc-table { width:100%; border-collapse:collapse; font-family:'JetBrains Mono',monospace; }
.nc-table th {
  font-size:.52rem; letter-spacing:2px; text-transform:uppercase;
  color:rgba(0,210,190,.5); padding:8px 12px; text-align:left;
  border-bottom:1px solid rgba(0,210,190,.1);
  background:rgba(4,10,18,.8);
}
.nc-table td {
  font-size:.7rem; padding:7px 12px; color:rgba(200,220,210,.75);
  border-bottom:1px solid rgba(255,255,255,.03);
  vertical-align:top;
}
.nc-table tr:hover td { background:rgba(0,210,190,.03); }
.nc-link {
  display:inline-block; font-size:.55rem; letter-spacing:1px;
  padding:3px 9px; border-radius:12px; text-decoration:none;
  transition:all .2s; white-space:nowrap;
}
.nc-link-v {
  background:rgba(0,210,190,.07); border:1px solid rgba(0,210,190,.2);
  color:#00d2be;
}
.nc-link-v:hover { background:rgba(0,210,190,.16); }
.nc-link-p {
  background:rgba(167,139,250,.07); border:1px solid rgba(167,139,250,.2);
  color:#a78bfa;
}
.nc-link-p:hover { background:rgba(167,139,250,.16); }
.nc-priority-high   { color:#22d88a; font-weight:600; }
.nc-priority-medium { color:#f0a020; }
.nc-priority-low    { color:#f87171; }
</style>"""


def render_new_columns(display_data: list[dict],
                        viz_base: str = _VIZ_BASE,
                        portal_base: str = _PORTAL_BASE,
                        max_rows: int = 50):
    """
    Render the 5 new smart columns as a supplemental HTML table.
    ADD-ONLY — placed after existing results, never replaces them.
    All operations wrapped in try/except.
    """
    if not _ST or not display_data:
        return
    try:
        st.markdown(_NEW_COL_CSS, unsafe_allow_html=True)
        st.markdown("""
<div style="font-family:'JetBrains Mono',monospace;font-size:.55rem;letter-spacing:3px;
  text-transform:uppercase;color:rgba(0,210,190,.4);margin:20px 0 8px">
  ⬡ Extended Intelligence Columns
</div>""", unsafe_allow_html=True)

        rows_html = ""
        shown = display_data[:max_rows]
        for c in shown:
            try:
                smiles   = c.get("SMILES") or c.get("smi") or c.get("smiles", "")
                cid      = c.get("ID") or c.get("compound_id") or smiles[:16] or "—"
                insight  = get_insight_summary(c)
                risk     = get_risk_summary(c)
                rec      = get_recommendation(c)
                priority = get_development_priority(c)
                p_class  = ("nc-priority-high" if "HIGH" in priority else
                            "nc-priority-medium" if "MEDIUM" in priority else
                            "nc-priority-low")

                viz_url    = get_visualization_url(smiles, viz_base) if smiles else "#"
                portal_url = get_portal_url(smiles, portal_base) if smiles else "#"

                viz_link = (f'<a href="{viz_url}" target="_blank" class="nc-link nc-link-v">'
                            f'🔬 Visualize</a>') if smiles else "—"
                portal_link = (f'<a href="{portal_url}" target="_blank" class="nc-link nc-link-p">'
                               f'📊 Portal</a>') if smiles else "—"

                rows_html += (
                    f'<tr>'
                    f'<td style="color:rgba(0,210,190,.8);font-weight:500">{cid[:18]}</td>'
                    f'<td>{viz_link}</td>'
                    f'<td>{portal_link}</td>'
                    f'<td style="max-width:220px">{insight}</td>'
                    f'<td style="max-width:160px">{risk}</td>'
                    f'<td style="max-width:200px">{rec}</td>'
                    f'<td class="{p_class}">{priority}</td>'
                    f'</tr>'
                )
            except Exception:
                continue

        table_html = f"""
<div style="background:rgba(4,10,18,.7);border:1px solid rgba(0,210,190,.08);
  border-radius:12px;overflow:hidden;margin:8px 0">
  <table class="nc-table">
    <thead>
      <tr>
        <th>Compound</th>
        <th>🔬 Visualization</th>
        <th>📊 Data Portal</th>
        <th>🧠 Insight</th>
        <th>⚠️ Risk Flags</th>
        <th>🎯 Recommendation</th>
        <th>Dev Priority</th>
      </tr>
    </thead>
    <tbody>
      {rows_html}
    </tbody>
  </table>
</div>"""

        if len(display_data) > max_rows:
            table_html += (
                f'<div style="font-family:\'JetBrains Mono\',monospace;font-size:.55rem;'
                f'color:rgba(0,210,190,.3);padding:6px 0">'
                f'Showing {max_rows} of {len(display_data)} compounds</div>')

        st.markdown(table_html, unsafe_allow_html=True)

    except Exception:
        pass  # never crash the page


# ─────────────────────────────────────────────────────────────────────────────
# DATAFRAME INJECTION — adds new columns to a pandas DataFrame in-place
# ─────────────────────────────────────────────────────────────────────────────

def inject_columns_into_df(df: "pd.DataFrame",
                             compounds: list[dict]) -> "pd.DataFrame":
    """
    Add the 5 new smart columns directly to a pandas DataFrame.
    Non-destructive: only appends new columns, never modifies existing ones.
    """
    if not _PD:
        return df
    try:
        rows = compounds[:len(df)]
        df = df.copy()
        df["🔬 Visualization"]   = [f'[View]({get_visualization_url(c.get("SMILES",""))})' for c in rows]
        df["📊 Data Portal"]     = [f'[Open]({get_portal_url(c.get("SMILES",""))})' for c in rows]
        df["🧠 Insight"]         = [get_insight_summary(c) for c in rows]
        df["⚠️ Risk Flags"]      = [get_risk_summary(c) for c in rows]
        df["🎯 Recommendation"]  = [get_recommendation(c) for c in rows]
        df["⬆ Dev Priority"]    = [get_development_priority(c) for c in rows]
        return df
    except Exception:
        return df


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR STATUS STRIP
# ─────────────────────────────────────────────────────────────────────────────

def render_sidebar_service_links():
    """
    Add visualization + portal quick-access links to the Streamlit sidebar.
    ADD-ONLY — call from sidebar section in app.py.
    """
    if not _ST:
        return
    try:
        st.sidebar.markdown("""
<div style="background:rgba(0,210,190,.04);border:1px solid rgba(0,210,190,.12);
  border-radius:8px;padding:12px 14px;margin:8px 0">
  <div style="font-family:'JetBrains Mono',monospace;font-size:.5rem;letter-spacing:2px;
    text-transform:uppercase;color:rgba(0,210,190,.4);margin-bottom:8px">
    ⬡ External Services
  </div>
  <a href="/visualization" target="_blank"
     style="display:block;font-size:.68rem;color:#00d2be;text-decoration:none;
     padding:4px 0;border-bottom:1px solid rgba(0,210,190,.07)">
    🔬 Visualization App →
  </a>
  <a href="/data_portal" target="_blank"
     style="display:block;font-size:.68rem;color:#a78bfa;text-decoration:none;
     padding:4px 0">
    📊 Data Portal →
  </a>
</div>""", unsafe_allow_html=True)
    except Exception:
        pass
