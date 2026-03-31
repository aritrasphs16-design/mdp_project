#!/usr/bin/env python3
"""
launch_all.py
────────────────────────────────────────────────────────────────────────────
ChemoFilter · Multi-Service Launcher
────────────────────────────────────────────────────────────────────────────

Starts all three Streamlit services simultaneously:
  Port 8501 — Main ChemoFilter App        (streamlit run app.py)
  Port 8502 — Visualization App           (streamlit run visualization_app/app.py)
  Port 8503 — Data Portal                 (streamlit run data_portal/app.py)

Usage:
  python launch_all.py          # start all 3
  python launch_all.py --main   # main app only
  python launch_all.py --viz    # visualization only
  python launch_all.py --portal # data portal only
────────────────────────────────────────────────────────────────────────────
"""
import subprocess
import sys
import os
import time
import signal

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

SERVICES = {
    "main":   {"script": "app.py",                       "port": 8501, "label": "Main App"},
    "viz":    {"script": "visualization_app/app.py",     "port": 8502, "label": "Visualization"},
    "portal": {"script": "data_portal/app.py",           "port": 8503, "label": "Data Portal"},
}

processes = []

def _streamlit_cmd(script: str, port: int) -> list[str]:
    return [
        sys.executable, "-m", "streamlit", "run",
        os.path.join(BASE_DIR, script),
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.runOnSave", "false",
        "--server.fileWatcherType", "none",
        "--browser.gatherUsageStats", "false",
    ]

def start_service(name: str):
    svc = SERVICES[name]
    cmd = _streamlit_cmd(svc["script"], svc["port"])
    print(f"  🚀 Starting {svc['label']} on port {svc['port']}…")
    proc = subprocess.Popen(cmd, cwd=BASE_DIR,
                             stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
    processes.append((name, proc))
    return proc

def stop_all():
    print("\n  🛑 Shutting down all services…")
    for name, proc in processes:
        try:
            proc.terminate()
            proc.wait(timeout=3)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
        print(f"     ✓ {SERVICES[name]['label']} stopped")

def signal_handler(sig, frame):
    stop_all()
    sys.exit(0)

if __name__ == "__main__":
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    args = sys.argv[1:]
    if "--main" in args:
        to_start = ["main"]
    elif "--viz" in args:
        to_start = ["viz"]
    elif "--portal" in args:
        to_start = ["portal"]
    else:
        to_start = ["main", "viz", "portal"]

    print("\n╔══════════════════════════════════════════════════╗")
    print("║  ⬡  ChemoFilter · Multi-Service Platform        ║")
    print("║     Crystalline Noir Edition · VIT Chennai 2026  ║")
    print("╚══════════════════════════════════════════════════╝\n")

    for name in to_start:
        start_service(name)
        time.sleep(0.5)

    print()
    print("  ✅ Services running:")
    for name in to_start:
        svc = SERVICES[name]
        print(f"     🌐 {svc['label']:<22} → http://localhost:{svc['port']}")

    print()
    print("  Press Ctrl+C to stop all services.\n")

    try:
        while True:
            # Check if any process died unexpectedly
            for name, proc in processes:
                if proc.poll() is not None:
                    print(f"  ⚠️  {SERVICES[name]['label']} exited (code {proc.returncode}). Restarting…")
                    new_proc = start_service(name)
                    processes[:] = [(n, p if n != name else new_proc) for n, p in processes]
            time.sleep(5)
    except KeyboardInterrupt:
        stop_all()
