"""Visual effect helpers (drop shadows, focus rings).

QSS does not support `box-shadow`, so subtle elevation is achieved with
``QGraphicsDropShadowEffect``. These helpers centralize the values so the
look stays consistent and theme-aware.
"""

from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

from .tokens import get_tokens


def _shadow_color(mode: str, strong: bool = False) -> QtGui.QColor:
    t = get_tokens(mode)
    raw = t["shadow_color_strong"] if strong else t["shadow_color"]
    color = QtGui.QColor()
    # QColor accepts CSS rgba() strings via setNamedColor on Qt 6
    color.setNamedColor(raw)
    if not color.isValid():
        # Fallback: parse rgba(r, g, b, a) manually
        try:
            inside = raw.strip().removeprefix("rgba(").removesuffix(")")
            parts = [p.strip() for p in inside.split(",")]
            r, g, b = (int(parts[0]), int(parts[1]), int(parts[2]))
            a = float(parts[3]) if len(parts) > 3 else 1.0
            color = QtGui.QColor(r, g, b, int(a * 255))
        except Exception:
            color = QtGui.QColor(0, 0, 0, 60)
    return color


def apply_card_shadow(widget: QtWidgets.QWidget, level: int = 1, mode: str = "light") -> None:
    """Apply a subtle drop shadow to a widget. ``level`` 1..3 = z-elevation."""
    if widget is None:
        return
    blurs = {1: 8, 2: 16, 3: 28}
    offsets = {1: 1, 2: 2, 3: 4}
    blur = blurs.get(level, 8)
    offset = offsets.get(level, 1)
    eff = QtWidgets.QGraphicsDropShadowEffect(widget)
    eff.setBlurRadius(blur)
    eff.setOffset(0, offset)
    eff.setColor(_shadow_color(mode, strong=level >= 3))
    widget.setGraphicsEffect(eff)


def clear_effect(widget: QtWidgets.QWidget) -> None:
    """Remove any QGraphicsEffect from a widget."""
    if widget is not None:
        widget.setGraphicsEffect(None)


__all__ = ["apply_card_shadow", "clear_effect"]
