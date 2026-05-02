"""Shared pytest fixtures.

Forces Qt's offscreen platform plugin so widget tests run without a display,
and provides a session-scoped QApplication. Qt rejects multiple QApplication
instances, hence the session scope.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Ensure repo root is on sys.path so `import gui` and project packages resolve
# regardless of where pytest is invoked from.
_REPO_ROOT = Path(__file__).resolve().parent.parent
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

import pytest
from PySide6 import QtWidgets


@pytest.fixture(scope="session")
def qapp() -> QtWidgets.QApplication:
    app = QtWidgets.QApplication.instance()
    if app is None:
        app = QtWidgets.QApplication(sys.argv)
    return app
