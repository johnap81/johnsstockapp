"""
Entry point for hosting (e.g. Render). The real app lives in server.py at the project root.

Render may set "Root Directory" to the repo root or to src/ — the dashboard start command
should be one of:
  - python src/server.py   (repo root)
  - python server.py       (root directory = src)
This file always resolves the parent server via __file__, not the process cwd.
"""
from __future__ import annotations

import os
import runpy
import sys


def _candidate_server_paths() -> list[str]:
    """Return possible locations for the real app entry `server.py`.

    Normal layout:
      repo/server.py
      repo/src/server.py  (this launcher)

    Common mistake when uploading via GitHub web UI:
      repo/<Some Folder>/server.py
      repo/<Some Folder>/src/server.py

    You can override with env SERVER_PY_PATH=/absolute/or/relative/path/server.py on Render.
    """
    here = os.path.dirname(os.path.abspath(__file__))
    override = ""
    try:
        # In some hosted runtimes, `os.environ.get()` can misbehave due to a broken env mapping.
        override = (os.environ.get("SERVER_PY_PATH") or "").strip()
    except RecursionError:
        override = ""
    candidates: list[str] = []
    if override:
        candidates.append(override if os.path.isabs(override) else os.path.abspath(override))

    # Same directory (works if Render Root Directory is mistakenly set to `src/`)
    candidates.append(os.path.join(here, "server.py"))
    # Repo root (normal)
    candidates.append(os.path.abspath(os.path.join(here, os.pardir, "server.py")))
    # One extra nesting level (upload-into-subfolder layouts)
    candidates.append(os.path.abspath(os.path.join(here, os.pardir, os.pardir, "server.py")))
    # CWD fallbacks (some hosts run with an unexpected working directory)
    cwd = os.path.abspath(os.getcwd())
    candidates.append(os.path.join(cwd, "server.py"))
    candidates.append(os.path.join(cwd, "..", "server.py"))

    # De-dupe while preserving order
    out: list[str] = []
    seen: set[str] = set()
    for p in candidates:
        ap = os.path.abspath(os.path.normpath(p))
        if ap not in seen:
            seen.add(ap)
            out.append(ap)
    return out


if __name__ == "__main__":
    tried = _candidate_server_paths()
    target = next((p for p in tried if os.path.isfile(p)), "")
    if not target:
        print(
            "FATAL: could not find the real `server.py` for this deploy.\n"
            "This launcher expects the main app file named exactly `server.py`.\n"
            "Tried:\n  - "
            + "\n  - ".join(tried)
            + "\n\nFix:\n"
            "- Best: ensure GitHub repo root contains BOTH `server.py` and `src/server.py` (same as your Mac folder).\n"
            "- Or set Render env var SERVER_PY_PATH to the deployed path of `server.py`.\n"
            "- Or set Render Root Directory to the folder that actually contains `server.py` and start with: `python server.py`\n",
            file=sys.stderr,
        )
        raise SystemExit(1)
    runpy.run_path(target, run_name="__main__")
