"""Keep this repository's generic ``src`` package isolated in shared dev envs."""

from __future__ import annotations

from pathlib import Path
import sys


REPO_ROOT = Path(__file__).resolve().parents[1]
root_text = str(REPO_ROOT)
if root_text in sys.path:
    sys.path.remove(root_text)
sys.path.insert(0, root_text)

loaded_src = sys.modules.get("src")
loaded_path = Path(getattr(loaded_src, "__file__", "") or "/").resolve()
if loaded_src is not None and REPO_ROOT not in loaded_path.parents:
    for module_name in tuple(sys.modules):
        if module_name == "src" or module_name.startswith("src."):
            sys.modules.pop(module_name, None)
