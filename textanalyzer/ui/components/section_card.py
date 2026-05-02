"""SectionCard — titled container that replaces ad-hoc QGroupBox usage."""

from __future__ import annotations

from typing import Sequence

from PySide6 import QtCore, QtWidgets


class SectionCard(QtWidgets.QFrame):
    """A card with a bold title, optional subtitle, body layout, and actions."""

    def __init__(
        self,
        title: str,
        *,
        subtitle: str = "",
        actions: Sequence[QtWidgets.QWidget] = (),
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("SectionCard")

        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 12)
        root.setSpacing(8)

        # Header row.
        header = QtWidgets.QHBoxLayout()
        header.setSpacing(8)
        self._title = QtWidgets.QLabel(title)
        title_font = self._title.font()
        title_font.setBold(True)
        self._title.setFont(title_font)
        header.addWidget(self._title, 1)
        for w in actions:
            header.addWidget(w)
        root.addLayout(header)

        if subtitle:
            sub = QtWidgets.QLabel(subtitle)
            sub.setProperty("role", "muted")
            sub.setWordWrap(True)
            root.addWidget(sub)

        # Placeholder for caller-provided content.
        self.body = QtWidgets.QVBoxLayout()
        self.body.setContentsMargins(0, 0, 0, 0)
        self.body.setSpacing(6)
        root.addLayout(self.body)

    def set_title(self, text: str) -> None:
        self._title.setText(text)


__all__ = ["SectionCard"]
