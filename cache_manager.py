"""
cache_manager.py
Deterministic caching layer for expensive computations.
Uses streamlit cache_data / cache_resource where available,
with a simple dict fallback for non-Streamlit contexts.
"""

import hashlib
import json
import time
from typing import Any, Callable, Optional

# ── Simple in-process fallback cache ────────────────────────────────────────
_MEMORY_CACHE: dict = {}
_CACHE_TTL: dict = {}           # key -> expiry timestamp (0 = no expiry)
_DEFAULT_TTL = 3600             # 1 hour


def _make_key(*args, **kwargs) -> str:
    """Create a deterministic cache key from arbitrary args."""
    try:
        raw = json.dumps({"args": args, "kwargs": kwargs}, sort_keys=True, default=str)
    except Exception:
        raw = str(args) + str(kwargs)
    return hashlib.md5(raw.encode()).hexdigest()


def cache_get(key: str) -> Optional[Any]:
    """Retrieve value from memory cache (None if missing/expired)."""
    expiry = _CACHE_TTL.get(key, 0)
    if expiry and time.time() > expiry:
        _MEMORY_CACHE.pop(key, None)
        _CACHE_TTL.pop(key, None)
        return None
    return _MEMORY_CACHE.get(key)


def cache_set(key: str, value: Any, ttl: int = _DEFAULT_TTL) -> None:
    """Store value in memory cache with optional TTL (seconds)."""
    _MEMORY_CACHE[key] = value
    _CACHE_TTL[key] = time.time() + ttl if ttl else 0


def cache_delete(key: str) -> None:
    _MEMORY_CACHE.pop(key, None)
    _CACHE_TTL.pop(key, None)


def cache_clear_all() -> int:
    """Clear entire cache. Returns number of entries cleared."""
    n = len(_MEMORY_CACHE)
    _MEMORY_CACHE.clear()
    _CACHE_TTL.clear()
    return n


def cached_call(fn: Callable, *args, ttl: int = _DEFAULT_TTL, **kwargs) -> Any:
    """
    Call fn(*args, **kwargs) with automatic caching.
    Uses fn.__name__ + args + kwargs as cache key.
    """
    key = _make_key(fn.__name__, *args, **kwargs)
    cached = cache_get(key)
    if cached is not None:
        return cached
    result = fn(*args, **kwargs)
    cache_set(key, result, ttl)
    return result


def get_cache_stats() -> dict:
    """Return cache usage statistics."""
    now = time.time()
    expired = sum(1 for k, exp in _CACHE_TTL.items() if exp and now > exp)
    return {
        "total_entries": len(_MEMORY_CACHE),
        "expired_entries": expired,
        "active_entries": len(_MEMORY_CACHE) - expired,
    }


# ── Streamlit-aware decorators ───────────────────────────────────────────────
def get_st_cached_resource(fn: Callable) -> Callable:
    """
    Wrap fn with st.cache_resource if Streamlit is available,
    otherwise return fn unchanged.
    """
    try:
        import streamlit as st
        return st.cache_resource(fn)
    except Exception:
        return fn


def get_st_cached_data(fn: Callable, ttl: int = _DEFAULT_TTL) -> Callable:
    """
    Wrap fn with st.cache_data(ttl=...) if Streamlit is available,
    otherwise return fn unchanged.
    """
    try:
        import streamlit as st
        return st.cache_data(ttl=ttl)(fn)
    except Exception:
        return fn
