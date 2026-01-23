import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from smartcatalog.main import start_ui

if __name__ == "__main__":
    start_ui()