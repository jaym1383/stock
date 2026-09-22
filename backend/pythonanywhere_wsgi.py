import sys
from pathlib import Path


PROJECT_DIR = Path.home() / "stock" / "backend"
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from app import app as application
