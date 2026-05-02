"""Backwards-compat shim for ``textanalyzer.engine.cluster``.

The real implementation moved into the package; this module is kept so
external scripts and legacy imports (``from cluster_tool import ...``,
``import cluster_tool as ct``, etc.) keep working.

Prefer the canonical path in new code:

    from textanalyzer.engine.cluster import cluster_texts
"""
from __future__ import annotations

import sys

from textanalyzer.engine import cluster as _impl

# Make this module name an alias for the real one. After this, every
# attribute lookup, including private/underscore-prefixed names like
# ``_HDBSCAN_AVAILABLE`` that legacy code touches, resolves through the
# canonical module object — no manual re-export list needed.
sys.modules[__name__] = _impl

# Allow `python cluster_tool.py ...` to keep working as a CLI entry point
# by delegating to the engine module's main().
if __name__ == "__main__":
    _impl.main()
