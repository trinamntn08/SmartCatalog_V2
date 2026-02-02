import sys
from pathlib import Path

def _project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

# add likely package roots to import path (dev + frozen)
root = _project_root()
candidate_roots = [
    root,
    root / "src",
]
for p in candidate_roots:
    if p.exists():
        sys.path.insert(0, str(p))

from smartcatalog.main import start_ui

if __name__ == "__main__":
    project_dir = _project_root()
    start_ui(project_dir=project_dir)
