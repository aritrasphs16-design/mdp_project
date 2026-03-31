"""
engine_orchestrator.py
Central controller that manages all engine execution, dispatching, fallbacks,
and unified metrics tracking.
"""

import time
import traceback
from typing import Any, Callable, Dict, Optional

# ── In-memory registry & metrics store ──────────────────────────────────────
_ENGINE_REGISTRY: Dict[str, Callable] = {}
_METRICS: Dict[str, list] = {}          # engine_name -> list of run dicts


def register_engine(name: str, factory: Callable) -> None:
    """Register an engine factory/callable under a logical name."""
    _ENGINE_REGISTRY[name] = factory
    if name not in _METRICS:
        _METRICS[name] = []


def get_registered_engines() -> list:
    """Return list of all registered engine names."""
    return list(_ENGINE_REGISTRY.keys())


def _record(engine_name: str, success: bool, elapsed: float, error: str = "") -> None:
    entry = {
        "engine": engine_name,
        "success": success,
        "elapsed_ms": round(elapsed * 1000, 2),
        "error": error,
        "timestamp": time.time(),
    }
    _METRICS.setdefault(engine_name, []).append(entry)
    # Keep only last 200 records per engine
    if len(_METRICS[engine_name]) > 200:
        _METRICS[engine_name] = _METRICS[engine_name][-200:]


def dispatch(engine_name: str, *args, fallback: Optional[str] = None, **kwargs) -> Dict[str, Any]:
    """
    Dispatch a task to the named engine.
    Returns dict: {result, engine_used, success, elapsed_ms, error}
    Falls back to `fallback` engine if primary fails.
    """
    result = _run_engine(engine_name, *args, **kwargs)
    if not result["success"] and fallback and fallback in _ENGINE_REGISTRY:
        fallback_result = _run_engine(fallback, *args, **kwargs)
        fallback_result["fallback_used"] = engine_name
        return fallback_result
    return result


def _run_engine(engine_name: str, *args, **kwargs) -> Dict[str, Any]:
    if engine_name not in _ENGINE_REGISTRY:
        return {
            "result": None,
            "engine_used": engine_name,
            "success": False,
            "elapsed_ms": 0,
            "error": f"Engine '{engine_name}' not registered.",
        }
    t0 = time.perf_counter()
    try:
        factory = _ENGINE_REGISTRY[engine_name]
        result = factory(*args, **kwargs)
        elapsed = time.perf_counter() - t0
        _record(engine_name, True, elapsed)
        return {
            "result": result,
            "engine_used": engine_name,
            "success": True,
            "elapsed_ms": round(elapsed * 1000, 2),
            "error": "",
        }
    except Exception as exc:
        elapsed = time.perf_counter() - t0
        err_msg = traceback.format_exc()
        _record(engine_name, False, elapsed, str(exc))
        return {
            "result": None,
            "engine_used": engine_name,
            "success": False,
            "elapsed_ms": round(elapsed * 1000, 2),
            "error": err_msg,
        }


def get_all_metrics() -> Dict[str, list]:
    """Return full metrics dict (copy)."""
    return {k: list(v) for k, v in _METRICS.items()}


def get_engine_summary() -> Dict[str, Dict]:
    """
    Return per-engine summary:
      {engine_name: {calls, successes, failures, avg_ms, last_error}}
    """
    summary = {}
    for name, runs in _METRICS.items():
        if not runs:
            summary[name] = {"calls": 0, "successes": 0, "failures": 0,
                             "avg_ms": 0.0, "last_error": ""}
            continue
        successes = sum(1 for r in runs if r["success"])
        failures = len(runs) - successes
        avg_ms = round(sum(r["elapsed_ms"] for r in runs) / len(runs), 2)
        last_err = next((r["error"] for r in reversed(runs) if r["error"]), "")
        summary[name] = {
            "calls": len(runs),
            "successes": successes,
            "failures": failures,
            "avg_ms": avg_ms,
            "last_error": last_err,
        }
    return summary
