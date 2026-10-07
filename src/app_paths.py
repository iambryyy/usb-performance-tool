import os
import sys


def app_dir():
    """
    Directory the running program actually lives in.

    A frozen PyInstaller --onefile .exe extracts itself into a fresh temp
    folder on every launch, so anything resolved from __file__ there (logs,
    saved router sessions, test files, results) would silently vanish
    between runs. Anchor to the .exe's own folder when frozen, and to this
    script's folder otherwise -- same result as today whenever this isn't
    running as a frozen build.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def resolve_path(path):
    """A config-supplied relative path (e.g. "test_files") anchored to
    app_dir() instead of whatever the current working directory happens to
    be -- so results land in the same place next to the program regardless
    of how it was launched. Absolute paths pass through unchanged."""
    return path if os.path.isabs(path) else os.path.join(app_dir(), path)
