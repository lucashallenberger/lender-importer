"""Per-user, writable locations for config + learned data.

In a packaged app the program folder is read-only, so config.ini and learned.json
live in the OS's per-user app-data folder instead:
  Windows: %APPDATA%\\LenderImporter
  macOS:   ~/Library/Application Support/LenderImporter
During development (running from source) we keep using files in the project folder.
"""

import os
import sys

APP_NAME = "LenderImporter"


def is_frozen():
    return getattr(sys, "frozen", False)


def user_data_dir():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    path = os.path.join(base, APP_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def project_dir():
    return os.path.dirname(os.path.abspath(__file__))


def data_file(name):
    """Where a data file (config.ini, learned.json) should live for reads/writes.
    Packaged app -> per-user folder. From source -> project folder."""
    return os.path.join(user_data_dir() if is_frozen() else project_dir(), name)
