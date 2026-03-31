"""
visualization.py
Visualization tools for ChemoFilter — reaction graphs, chemical networks,
analysis charts, performance metrics. All Streamlit-compatible (Plotly).
"""

from typing import Dict, List, Optional
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd


# ── Color palette (matches existing UI theme) ────────────────────────────────
_COLORS = {
    "primary": "#7C3AED",
    "success": "#10B981",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "info": "#3B82F6",
    "muted": "#6B7280",
    "bg": "#0F172A",
    "surface": "#1E293B",
}

_PLOTLY_TEMPLATE = dict(
    layout=dict(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E2E8F0", family="Inter, sans-serif"),
        margin=dict(l=40, r=20, t=40, b=40),
    )
)


def radar_chart(
    categories: List[str],
    values: List[float],
    title: str = "Property Radar",
    max_val: float = 1.0,
) -> go.Figure:
    """Radar/spider chart for multi-property display."""
    cats = categories + [categories[0]]
    vals = values + [values[0]]
    fig = go.Figure(go.Scatterpolar(
        r=vals, theta=cats,
        fill="toself",
        line_color=_COLORS["primary"],
        fillcolor=f"rgba(124,58,237,0.25)",
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=16)),
        polar=dict(
            bgcolor=_COLORS["surface"],
            radialaxis=dict(visible=True, range=[0, max_val],
                            gridcolor="#334155", tickfont=dict(size=9)),
            angularaxis=dict(gridcolor="#334155"),
        ),
        showlegend=False,
        **_PLOTLY_TEMPLATE["layout"],
    )
    return fig


def bar_chart(
    labels: List[str],
    values: List[float],
    title: str = "",
    color_scale: bool = True,
    threshold: Optional[float] = None,
) -> go.Figure:
    """Horizontal bar chart with optional threshold line."""
    colors = []
    for v in values:
        if threshold is not None:
            colors.append(_COLORS["success"] if v <= threshold else _COLORS["danger"])
        else:
            colors.append(_COLORS["primary"])
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=colors,
        text=[f"{v:.2f}" for v in values],
        textposition="outside",
    ))
    if threshold is not None:
        fig.add_vline(x=threshold, line_dash="dash",
                      line_color=_COLORS["warning"], annotation_text=f"Threshold={threshold}")
    fig.update_layout(title=dict(text=title, font=dict(size=15)),
                      xaxis=dict(gridcolor="#334155"),
                      yaxis=dict(gridcolor="#334155"),
                      **_PLOTLY_TEMPLATE["layout"])
    return fig


def gauge_chart(value: float, title: str = "", min_val: float = 0,
                max_val: float = 1, thresholds: Optional[List[float]] = None) -> go.Figure:
    """Gauge chart for single score display (0–1 or custom range)."""
    thresholds = thresholds or [max_val * 0.33, max_val * 0.66]
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=value,
        title={"text": title, "font": {"size": 16, "color": "#E2E8F0"}},
        gauge={
            "axis": {"range": [min_val, max_val], "tickcolor": "#94A3B8"},
            "bar": {"color": _COLORS["primary"]},
            "steps": [
                {"range": [min_val, thresholds[0]], "color": "rgba(239,68,68,0.3)"},
                {"range": [thresholds[0], thresholds[1]], "color": "rgba(245,158,11,0.3)"},
                {"range": [thresholds[1], max_val], "color": "rgba(16,185,129,0.3)"},
            ],
            "bgcolor": _COLORS["surface"],
            "bordercolor": "#334155",
        },
        number={"font": {"color": "#E2E8F0"}},
    ))
    fig.update_layout(**_PLOTLY_TEMPLATE["layout"])
    return fig


def performance_timeline(events: List[dict], title: str = "Engine Performance Timeline") -> go.Figure:
    """
    Line chart of elapsed_ms over time for engine performance events.
    events: list of dicts with keys: ts, elapsed_ms, label, success
    """
    if not events:
        fig = go.Figure()
        fig.update_layout(title=title, **_PLOTLY_TEMPLATE["layout"])
        return fig
    df = pd.DataFrame(events)
    df["time"] = pd.to_datetime(df["ts"], unit="s")
    df["color"] = df["success"].map({True: _COLORS["success"], False: _COLORS["danger"]})
    labels = df["label"].unique().tolist()
    fig = go.Figure()
    for lbl in labels:
        sub = df[df["label"] == lbl]
        fig.add_trace(go.Scatter(
            x=sub["time"], y=sub["elapsed_ms"],
            mode="lines+markers",
            name=lbl,
            marker=dict(size=6),
            line=dict(width=2),
        ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=15)),
        xaxis=dict(title="Time", gridcolor="#334155"),
        yaxis=dict(title="Latency (ms)", gridcolor="#334155"),
        **_PLOTLY_TEMPLATE["layout"],
    )
    return fig


def engine_stats_bar(summary: dict, title: str = "Engine Call Statistics") -> go.Figure:
    """
    Bar chart from engine_orchestrator.get_engine_summary() output.
    summary: {engine_name: {calls, successes, failures, avg_ms}}
    """
    names = list(summary.keys())
    calls = [summary[n].get("calls", 0) for n in names]
    successes = [summary[n].get("successes", 0) for n in names]
    failures = [summary[n].get("failures", 0) for n in names]
    fig = go.Figure(data=[
        go.Bar(name="Successes", x=names, y=successes, marker_color=_COLORS["success"]),
        go.Bar(name="Failures", x=names, y=failures, marker_color=_COLORS["danger"]),
    ])
    fig.update_layout(
        barmode="stack",
        title=dict(text=title, font=dict(size=15)),
        xaxis=dict(gridcolor="#334155"),
        yaxis=dict(title="Calls", gridcolor="#334155"),
        **_PLOTLY_TEMPLATE["layout"],
    )
    return fig


def property_heatmap(data: Dict[str, Dict[str, float]], title: str = "Property Heatmap") -> go.Figure:
    """
    2D heatmap for comparing multiple compounds vs multiple properties.
    data: {compound_name: {property_name: value}}
    """
    compounds = list(data.keys())
    if not compounds:
        return go.Figure()
    props = list(data[compounds[0]].keys())
    z = [[data[c].get(p, 0) for p in props] for c in compounds]
    fig = go.Figure(go.Heatmap(
        z=z, x=props, y=compounds,
        colorscale="Viridis",
        hoverongaps=False,
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=15)),
        **_PLOTLY_TEMPLATE["layout"],
    )
    return fig
