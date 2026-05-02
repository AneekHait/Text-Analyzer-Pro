"""KBDChip — keyboard shortcut badge (e.g., ``Ctrl+R``)."""

from __future__ import annotations

from PySide6 import QtWidgets


class KBDChip(QtWidgets.QLabel):
    """Styled label that renders as a keyboard-key chip."""

    def __init__(self, text: str, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setProperty("role", "kbd")


__all__ = ["KBDChip"]
