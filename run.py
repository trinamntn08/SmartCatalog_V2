import sys
import os
from pathlib import Path

# add src/ to import path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from smartcatalog.main import start_ui

if __name__ == "__main__":
    project_dir = Path(__file__).resolve().parent
    start_ui(project_dir=project_dir)
