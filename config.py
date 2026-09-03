"""File paths and constants used across the app."""

import os
import sys


def _base_dir():
    """Folder the .exe (or script) lives in — works both frozen and unfrozen."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = _base_dir()
ENGINEERING = os.path.join(BASE_DIR, "engineering_database.xlsx")
CONFIG = os.path.join(BASE_DIR, "utility_config.xlsx")
DB_FILE = os.path.join(BASE_DIR, "verification.db")
