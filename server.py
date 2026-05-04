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


def _main_server_path() -> str:
    return os.path.abspath(
        os.path.join(os.path.dirname(__file__), os.pardir, "server.py")
    )


if __name__ == "__main__":
    target = _main_server_path()
    if not os.path.isfile(target):
        print(
            "FATAL: project root server.py is missing from this deploy.\n"
            f"Expected at: {target}\n"
            "Fix on your Mac: open Terminal in the project folder, then run:\n"
            "  git add server.py src/server.py && git commit -m 'Add server for Render' && git push\n"
            "Redeploy on Render only after GitHub shows server.py in the repository.",
            file=sys.stderr,
        )
        raise SystemExit(1)
    runpy.run_path(target, run_name="__main__")
