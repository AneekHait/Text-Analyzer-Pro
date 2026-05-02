"""Backwards-compat shim for ``textanalyzer.settings``.

Prefer the canonical path in new code:

    from textanalyzer import settings as app_settings
"""
from __future__ import annotations

import sys

from textanalyzer import settings as _impl

sys.modules[__name__] = _impl
