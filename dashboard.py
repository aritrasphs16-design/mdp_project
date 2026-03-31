"""
dashboard.py
New UI layer — extends app.py with:
  - Sidebar navigation & engine selection panel
  - Analytics dashboard (performance metrics)
  - Universal search panel
  - Execution log viewer
  - Visual result display helpers
  - Debug mode toggle

Import and call render_dashboard_sidebar() + render_analytics_tab() from app.py
or any Streamlit entry point. All additions are additive — nothing in existing
modules is modified.
"""

import streamlit as st
import time
from typing import Optional

# ── Lazy imports (avoid crashing if new modules not yet on path) ─────────────
def _safe_import(name):
    try:
        import importlib
        return importlib.import_module(name)
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR: Brand + Navigation + Engine Selection
# ─────────────────────────────────────────────────────────────────────────────

def render_dashboard_sidebar():
    """
    Call this from the main app after existing sidebar content.
    Adds: Engine Selection Panel, Search toggle, Debug toggle.
    """
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 🧬 Engine Control")

        engines = [
            "Auto (Smart Select)",
            "aether_engine_v10000",
            "quantum_accuracy_engine",
            "omnipotent_engine_v200",
            "nova_engine_v3000",
            "xenon_engine_v5000",
            "celestial_engine_v1000",
            "omega_engine_v2000",
        ]
        selected_engine = st.selectbox(
            "Active Engine",
            engines,
            index=0,
            key="dashboard_engine_select",
            help="Choose analysis engine. Auto selects based on query type.",
        )
        st.session_state["active_engine"] = selected_engine

        st.markdown("---")
        st.markdown("### 🔍 Universal Search")
        search_query = st.text_input(
            "Search databases",
            placeholder="drug name, scaffold, SMILES...",
            key="dashboard_search_input",
        )
        if st.button("Search All DBs", key="dashboard_search_btn", width='stretch'):
            if search_query.strip():
                st.session_state["dashboard_search_query"] = search_query.strip()
                st.session_state["show_search_results"] = True

        st.markdown("---")
        st.markdown("### ⚙️ Dev Tools")
        debug_mode = st.toggle("Debug Mode", value=st.session_state.get("debug_mode", False),
                               key="debug_mode_toggle")
        st.session_state["debug_mode"] = debug_mode

        perf_mode = st.toggle("Performance Monitor", value=st.session_state.get("perf_mode", False),
                              key="perf_mode_toggle")
        st.session_state["perf_mode"] = perf_mode

        if st.button("Clear Cache", key="clear_cache_btn", width='stretch'):
            cm = _safe_import("cache_manager")
            if cm:
                n = cm.cache_clear_all()
                st.toast(f"✅ Cache cleared ({n} entries)", icon="🗑️")
            else:
                st.cache_data.clear()
                st.toast("✅ Streamlit cache cleared")


# ─────────────────────────────────────────────────────────────────────────────
# SEARCH RESULTS PANEL
# ─────────────────────────────────────────────────────────────────────────────

def render_search_results():
    """Render search results if a search was triggered via sidebar."""
    if not st.session_state.get("show_search_results"):
        return
    query = st.session_state.get("dashboard_search_query", "")
    if not query:
        return

    se = _safe_import("search_engine")
    if se is None:
        st.warning("search_engine module not available.")
        return

    st.markdown(f"## 🔍 Search Results for: `{query}`")
    with st.spinner("Searching all databases..."):
        results = se.search_all(query, top_n=8, threshold=0.25)
        summary = se.get_search_summary(results)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Hits", summary["total_hits"])
    col2.metric("Drug Atlas Hits", summary["by_source"].get("drug_atlas", 0))
    col3.metric("Chemical DB Hits", summary["by_source"].get("chemical_db", 0))

    for source, hits in results.items():
        if hits:
            with st.expander(f"📦 {source.replace('_', ' ').title()} ({len(hits)} hits)", expanded=True):
                import pandas as pd
                rows = [{"Name": h["name"],
                         "Match Score": f"{h['score']:.0%}",
                         "Source": h["source"]} for h in hits]
                st.dataframe(pd.DataFrame(rows), use_container_width=True)

    if st.button("✕ Close Search Results", key="close_search"):
        st.session_state["show_search_results"] = False
        st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ANALYTICS DASHBOARD TAB
# ─────────────────────────────────────────────────────────────────────────────

def render_analytics_tab():
    """
    Renders a full analytics dashboard. Call inside a st.tab() block.
    Example usage in app.py:
        tab_analytics, = st.tabs(["📊 Analytics"])
        with tab_analytics:
            from dashboard import render_analytics_tab
            render_analytics_tab()
    """
    st.markdown("## 📊 System Analytics Dashboard")

    pm = _safe_import("performance_monitor")
    eo = _safe_import("engine_orchestrator")
    se_mod = _safe_import("safe_execution")
    cm = _safe_import("cache_manager")
    viz = _safe_import("visualization")

    # ── System Snapshot ──────────────────────────────────────────────────────
    st.markdown("### 🖥️ System Snapshot")
    c1, c2, c3, c4 = st.columns(4)
    if pm:
        sys_info = pm.get_system_summary()
        c1.metric("Total Calls Tracked", sys_info.get("total_calls", 0))
        c2.metric("Avg Success Rate", f"{sys_info.get('avg_success_rate', 100)}%")
        mem = sys_info.get("memory_mb", -1)
        c3.metric("Memory Usage", f"{mem} MB" if mem >= 0 else "N/A (psutil missing)")
        c4.metric("Event Log Size", sys_info.get("event_log_size", 0))
    else:
        c1.info("performance_monitor not loaded")

    if cm:
        stats = cm.get_cache_stats()
        st.markdown("#### Cache Status")
        cc1, cc2, cc3 = st.columns(3)
        cc1.metric("Cache Entries", stats["total_entries"])
        cc2.metric("Active", stats["active_entries"])
        cc3.metric("Expired", stats["expired_entries"])

    # ── Engine Stats ─────────────────────────────────────────────────────────
    st.markdown("### ⚡ Engine Statistics")
    if eo:
        summary = eo.get_engine_summary()
        if summary:
            if viz:
                fig = viz.engine_stats_bar(summary)
                st.plotly_chart(fig, width='stretch')
            import pandas as pd
            rows = [
                {
                    "Engine": name,
                    "Calls": s["calls"],
                    "Successes": s["successes"],
                    "Failures": s["failures"],
                    "Avg ms": s["avg_ms"],
                }
                for name, s in summary.items()
            ]
            if rows:
                st.dataframe(pd.DataFrame(rows), use_container_width=True)
        else:
            st.info("No engine calls recorded yet in this session.")
    else:
        st.info("engine_orchestrator not loaded")

    # ── Performance Timeline ─────────────────────────────────────────────────
    if pm and viz:
        st.markdown("### 📈 Performance Timeline")
        events = pm.get_recent_events(100)
        if events:
            fig = viz.performance_timeline(events)
            st.plotly_chart(fig, width='stretch')
        else:
            st.info("No performance events recorded yet.")

    # ── Execution Log ────────────────────────────────────────────────────────
    if se_mod:
        st.markdown("### 📋 Execution Log")
        exec_stats = se_mod.get_exec_stats()
        es1, es2, es3, es4 = st.columns(4)
        es1.metric("Total Executions", exec_stats["total"])
        es2.metric("Successes", exec_stats["successes"])
        es3.metric("Failures", exec_stats["failures"])
        es4.metric("Avg Latency", f"{exec_stats['avg_ms']} ms")

        log = se_mod.get_exec_log(last_n=30)
        if log:
            import pandas as pd
            rows = [
                {
                    "Function": e["fn"],
                    "Status": "✅" if e["success"] else "❌",
                    "Latency ms": e["elapsed_ms"],
                    "Fallback": "↩" if e.get("fallback") else "",
                    "Error": e["error"][:60] + "..." if e.get("error") and len(e.get("error","")) > 60 else e.get("error",""),
                }
                for e in reversed(log)
            ]
            st.dataframe(pd.DataFrame(rows), use_container_width=True)


# ─────────────────────────────────────────────────────────────────────────────
# DEBUG PANEL
# ─────────────────────────────────────────────────────────────────────────────

def render_debug_panel():
    """Render debug info panel. Only shows if debug_mode is True."""
    if not st.session_state.get("debug_mode"):
        return

    with st.expander("🐛 Debug Panel", expanded=False):
        st.markdown("**Active Engine:** " + st.session_state.get("active_engine", "None"))
        st.markdown("**Session State Keys:** " + ", ".join(sorted(st.session_state.keys())))

        router = _safe_import("router")
        if router:
            last_query = st.session_state.get("last_analysis_query", "")
            if last_query:
                explanation = router.explain_routing(last_query)
                st.json(explanation)

        cm = _safe_import("cache_manager")
        if cm:
            st.json(cm.get_cache_stats())


# ─────────────────────────────────────────────────────────────────────────────
# LOADING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def show_loading(message: str = "Analyzing..."):
    """Context manager: show spinner + progress bar while block executes."""
    return st.spinner(message)


def render_progress(label: str, steps: list):
    """
    Render an animated progress sequence.
    steps: list of step-label strings
    """
    bar = st.progress(0, text=label)
    for i, step in enumerate(steps, 1):
        bar.progress(int(i / len(steps) * 100), text=step)
        time.sleep(0.15)
    bar.empty()


# ─────────────────────────────────────────────────────────────────────────────
# ROUTING EXPLAINER WIDGET
# ─────────────────────────────────────────────────────────────────────────────

def render_routing_explainer(query: str):
    """Show which engine was auto-selected for the query (debug aid)."""
    if not st.session_state.get("debug_mode"):
        return
    router = _safe_import("router")
    if router is None:
        return
    expl = router.explain_routing(query)
    st.caption(
        f"🔀 Auto-routed to **{expl['selected_engine']}** "
        f"| keywords: {', '.join(expl['matched_keywords']) or 'none'}"
    )
