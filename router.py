"""
router.py
Unified routing entry point — combines engine_selector + engine_orchestrator
for a single call-site API used by UI and other modules.
"""

import engine_selector as es
import engine_orchestrator as eo
from typing import Any, Dict, Optional


def route(
    query: str,
    *args,
    context: Optional[str] = None,
    fallback: Optional[str] = None,
    **kwargs,
) -> Dict[str, Any]:
    """
    Auto-select engine based on query, dispatch, and return result dict.
    Result dict keys: result, engine_used, success, elapsed_ms, error, selected_by
    """
    engine_name = es.select_engine(query, context)
    result = eo.dispatch(engine_name, *args, fallback=fallback, **kwargs)
    result["selected_by"] = "auto" if not context else context
    result["query"] = query
    return result


def route_to(
    engine_name: str,
    *args,
    fallback: Optional[str] = None,
    **kwargs,
) -> Dict[str, Any]:
    """
    Explicitly route to a named engine with optional fallback.
    """
    result = eo.dispatch(engine_name, *args, fallback=fallback, **kwargs)
    result["selected_by"] = "manual"
    return result


def explain_routing(query: str) -> Dict[str, Any]:
    """
    Return a debug explanation of how a query would be routed,
    including top-3 candidate engines and selected engine.
    """
    primary = es.select_engine(query)
    candidates = es.select_engines_multi(query, top_n=3)
    table = es.get_routing_table()
    matched_keywords = [kw for kw, eng in table.items() if kw in query.lower()]
    return {
        "query": query,
        "selected_engine": primary,
        "candidate_engines": candidates,
        "matched_keywords": matched_keywords,
    }
