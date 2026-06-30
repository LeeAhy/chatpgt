from __future__ import annotations

import sys
from pathlib import Path


BUNDLED_PYTHON_PACKAGES = Path(
    "/Users/chandelar/.cache/codex-runtimes/codex-primary-runtime/dependencies/python"
)


def ensure_bundled_python_path() -> None:
    if not BUNDLED_PYTHON_PACKAGES.exists():
        return

    candidates = [BUNDLED_PYTHON_PACKAGES]
    candidates.extend(sorted(BUNDLED_PYTHON_PACKAGES.glob("lib/python*/site-packages")))

    for candidate in reversed(candidates):
        path = str(candidate)
        if path not in sys.path:
            sys.path.insert(0, path)
