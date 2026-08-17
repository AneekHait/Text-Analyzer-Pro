"""Generated icon assets used by QSS rules.

QSS `image: url(...)` needs a real file path. We bake small icons (just the
combobox chevron right now) at theme-apply time so the color matches the
current theme tokens, and cache them on disk under the user's temp dir.
"""
from __future__ import annotations

import os
import tempfile

from PySide6 import QtCore, QtGui


_CACHE_DIR = os.path.join(tempfile.gettempdir(), "text_analyzer_pro_assets")
_CACHE: dict[str, str] = {}


def _ensure_cache_dir() -> str:
    os.makedirs(_CACHE_DIR, exist_ok=True)
    return _CACHE_DIR


def chevron_down_url(color_hex: str, size: int = 10) -> str:
    """Render a downward chevron of the given color and return a URL-safe path.

    Cached per (color, size). The returned string is a forward-slash path,
    which Qt accepts inside `url("...")` on Windows without escaping.
    """
    key = f"chevron-{color_hex.lstrip('#')}-{size}.png"
    if key in _CACHE:
        return _CACHE[key]

    path = os.path.join(_ensure_cache_dir(), key)
    if not os.path.exists(path):
        pixmap = QtGui.QPixmap(size * 2, size * 2)
        pixmap.fill(QtCore.Qt.transparent)
        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        pen = QtGui.QPen(QtGui.QColor(color_hex))
        pen.setWidth(2)
        pen.setCapStyle(QtCore.Qt.RoundCap)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        painter.setPen(pen)
        # Draw a 'v' chevron centered in the pixmap.
        margin = size * 0.5
        top = size * 0.7
        bottom = size * 1.3
        left = margin
        right = size * 2 - margin
        mid = size
        painter.drawLine(QtCore.QPointF(left, top), QtCore.QPointF(mid, bottom))
        painter.drawLine(QtCore.QPointF(mid, bottom), QtCore.QPointF(right, top))
        painter.end()
        pixmap.save(path, "PNG")

    url = path.replace(os.sep, "/")
    _CACHE[key] = url
    return url


def checkmark_url(color_hex: str, size: int = 14) -> str:
    """Render a checkmark/tick of the given color and return a URL-safe path."""
    key = f"checkmark-{color_hex.lstrip('#')}-{size}.png"
    if key in _CACHE:
        return _CACHE[key]

    path = os.path.join(_ensure_cache_dir(), key)
    if not os.path.exists(path):
        pixmap = QtGui.QPixmap(size, size)
        pixmap.fill(QtCore.Qt.transparent)
        painter = QtGui.QPainter(pixmap)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        pen = QtGui.QPen(QtGui.QColor(color_hex))
        pen.setWidth(max(2, size // 7))
        pen.setCapStyle(QtCore.Qt.RoundCap)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        painter.setPen(pen)
        m = size * 0.15
        painter.drawLine(
            QtCore.QPointF(m, size * 0.55),
            QtCore.QPointF(size * 0.38, size * 0.78),
        )
        painter.drawLine(
            QtCore.QPointF(size * 0.38, size * 0.78),
            QtCore.QPointF(size - m, size * 0.25),
        )
        painter.end()
        pixmap.save(path, "PNG")

    url = path.replace(os.sep, "/")
    _CACHE[key] = url
    return url
