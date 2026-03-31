"""
performance_monitor.py
Lightweight performance tracking: runtime, latency, success rate, memory usage.
All data stored in-memory. No external dependencies.
"""

import time
import threading
from collections import deque
from typing import Dict, List, Optional
from contextlib import contextmanager

try:
    import psutil as _psutil
    _PSUTIL = True
except ImportError:
    _psutil = None
    _PSUTIL = False

# ── Storage ──────────────────────────────────────────────────────────────────
_lock = threading.Lock()
_timings: Dict[str, deque] = {}             # label -> bounded deque of ms values
_events: deque = deque(maxlen=1000)         # O(1) bounded append — replaces list.pop(0)
_EVENTS_LIMIT = 1000


def _push_event(label: str, elapsed_ms: float, success: bool, note: str = "") -> None:
    with _lock:
        entry = {
            "label": label,
            "elapsed_ms": elapsed_ms,
            "success": success,
            "note": note,
            "ts": time.time(),
        }
        _events.append(entry)           # deque auto-drops oldest when maxlen reached
        if label not in _timings:
            _timings[label] = deque(maxlen=500)
        _timings[label].append(elapsed_ms)


@contextmanager
def track(label: str, note: str = ""):
    """
    Context manager for timing a block of code.
    Usage:
        with performance_monitor.track("engine:aether"):
            result = aether.analyze()
    """
    t0 = time.perf_counter()
    success = True
    try:
        yield
    except Exception:
        success = False
        raise
    finally:
        elapsed = round((time.perf_counter() - t0) * 1000, 2)
        _push_event(label, elapsed, success, note)


def record(label: str, elapsed_ms: float, success: bool = True, note: str = "") -> None:
    """Manually record a timing event."""
    _push_event(label, elapsed_ms, success, note)


def get_stats(label: Optional[str] = None) -> dict:
    """
    Return stats for one label or all labels.
    Each entry: {count, avg_ms, min_ms, max_ms, p95_ms, success_rate}
    """
    with _lock:
        labels = [label] if label else list(_timings.keys())
        result = {}
        for lbl in labels:
            times = _timings.get(lbl, [])
            if not times:
                result[lbl] = {"count": 0}
                continue
            sorted_t = sorted(times)
            n = len(sorted_t)
            p95_idx = max(0, int(n * 0.95) - 1)
            successes = sum(1 for e in _events if e["label"] == lbl and e["success"])
            total_lbl = sum(1 for e in _events if e["label"] == lbl)
            result[lbl] = {
                "count": n,
                "avg_ms": round(sum(times) / n, 2),
                "min_ms": round(sorted_t[0], 2),
                "max_ms": round(sorted_t[-1], 2),
                "p95_ms": round(sorted_t[p95_idx], 2),
                "success_rate": round(successes / total_lbl * 100, 1) if total_lbl else 100.0,
            }
        return result


def get_recent_events(last_n: int = 50) -> list:
    with _lock:
        # deque doesn't support negative slicing — convert to list first
        ev = list(_events)
        return ev[-last_n:] if len(ev) > last_n else ev


def get_memory_usage_mb() -> float:
    """Return current process memory usage in MB (requires psutil)."""
    if not _PSUTIL:
        return -1.0
    try:
        import os
        proc = _psutil.Process(os.getpid())
        return round(proc.memory_info().rss / 1024 / 1024, 2)
    except Exception:
        return -1.0


def get_system_summary() -> dict:
    """High-level system health snapshot."""
    all_stats = get_stats()
    total_calls = sum(s.get("count", 0) for s in all_stats.values())
    avg_success = (
        round(sum(s.get("success_rate", 100) for s in all_stats.values()) / len(all_stats), 1)
        if all_stats else 100.0
    )
    return {
        "tracked_labels": len(all_stats),
        "total_calls": total_calls,
        "avg_success_rate": avg_success,
        "memory_mb": get_memory_usage_mb(),
        "event_log_size": len(_events),
    }


def reset() -> None:
    """Clear all stored performance data."""
    with _lock:
        _timings.clear()
        _events.clear()
