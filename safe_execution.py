"""
safe_execution.py
Controlled execution wrappers: try → catch → fallback → log.
Prevents unhandled exceptions from crashing the Streamlit app.
"""

import traceback
import time
import logging
from typing import Any, Callable, Dict, Optional, Tuple

# ── Module-level logger ──────────────────────────────────────────────────────
logger = logging.getLogger("safe_execution")
if not logger.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)s %(name)s: %(message)s",
                                      datefmt="%H:%M:%S"))
    logger.addHandler(_h)
    logger.setLevel(logging.DEBUG)

# ── Execution log (in-memory, last 500 entries) ──────────────────────────────
_EXEC_LOG: list = []
_LOG_LIMIT = 500


def _log_entry(fn_name: str, success: bool, elapsed_ms: float,
               error: str = "", fallback_used: bool = False) -> None:
    entry = {
        "fn": fn_name,
        "success": success,
        "elapsed_ms": elapsed_ms,
        "error": error,
        "fallback": fallback_used,
        "ts": time.time(),
    }
    _EXEC_LOG.append(entry)
    if len(_EXEC_LOG) > _LOG_LIMIT:
        _EXEC_LOG.pop(0)


def safe_run(
    fn: Callable,
    *args,
    fallback_fn: Optional[Callable] = None,
    default: Any = None,
    label: str = "",
    **kwargs,
) -> Tuple[Any, Dict]:
    """
    Execute fn(*args, **kwargs) safely.
    Returns (result, meta_dict).
    meta_dict keys: success, elapsed_ms, error, fallback_used
    """
    name = label or getattr(fn, "__name__", "anonymous")
    t0 = time.perf_counter()
    try:
        result = fn(*args, **kwargs)
        elapsed = round((time.perf_counter() - t0) * 1000, 2)
        _log_entry(name, True, elapsed)
        logger.debug(f"✓ {name} completed in {elapsed}ms")
        return result, {"success": True, "elapsed_ms": elapsed, "error": "", "fallback_used": False}
    except Exception as exc:
        elapsed = round((time.perf_counter() - t0) * 1000, 2)
        err_msg = traceback.format_exc()
        logger.warning(f"✗ {name} failed in {elapsed}ms: {exc}")
        _log_entry(name, False, elapsed, str(exc))

        # Try fallback
        if fallback_fn is not None:
            try:
                t1 = time.perf_counter()
                result = fallback_fn(*args, **kwargs)
                elapsed2 = round((time.perf_counter() - t1) * 1000, 2)
                fb_name = getattr(fallback_fn, "__name__", "fallback")
                _log_entry(fb_name, True, elapsed2, fallback_used=True)
                logger.info(f"↩ Fallback {fb_name} succeeded in {elapsed2}ms")
                return result, {
                    "success": True, "elapsed_ms": elapsed2,
                    "error": f"Primary failed: {exc}", "fallback_used": True,
                }
            except Exception as fb_exc:
                fb_elapsed = round((time.perf_counter() - t0) * 1000, 2)
                _log_entry("fallback", False, fb_elapsed, str(fb_exc), fallback_used=True)
                logger.error(f"✗ Fallback also failed: {fb_exc}")

        return default, {
            "success": False,
            "elapsed_ms": elapsed,
            "error": err_msg,
            "fallback_used": False,
        }


def safe_run_simple(fn: Callable, *args, default: Any = None, **kwargs) -> Any:
    """Simplified safe_run — returns result only (or default on failure)."""
    result, _ = safe_run(fn, *args, default=default, **kwargs)
    return result


def get_exec_log(last_n: int = 50) -> list:
    """Return last N execution log entries."""
    return _EXEC_LOG[-last_n:]


def get_exec_stats() -> Dict:
    """Summarize execution log: total, successes, failures, avg_ms."""
    if not _EXEC_LOG:
        return {"total": 0, "successes": 0, "failures": 0, "avg_ms": 0.0}
    successes = sum(1 for e in _EXEC_LOG if e["success"])
    avg_ms = round(sum(e["elapsed_ms"] for e in _EXEC_LOG) / len(_EXEC_LOG), 2)
    return {
        "total": len(_EXEC_LOG),
        "successes": successes,
        "failures": len(_EXEC_LOG) - successes,
        "avg_ms": avg_ms,
    }
