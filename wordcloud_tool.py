"""Backwards-compat shim for ``textanalyzer.engine.wordcloud``.

Prefer the canonical path in new code:

    from textanalyzer.engine.wordcloud import render_wordcloud
"""
from __future__ import annotations

import sys

from textanalyzer.engine import wordcloud as _impl

sys.modules[__name__] = _impl
