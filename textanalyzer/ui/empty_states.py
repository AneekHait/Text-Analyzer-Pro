"""Reusable empty-state and error-card widgets.

Both render as a card with an icon, a title, a body paragraph, and
optional action buttons. They use the ``QFrame#EmptyState`` /
``QFrame#ErrorCard`` QSS rules introduced in Phase 1.
"""

from __future__ import annotations

from typing import Callable

from PySide6 import QtCore, QtGui, QtWidgets

from ..utils.icons import _HAS_QTA, _icon


class EmptyState(QtWidgets.QWidget):
    """Centered empty-state card with icon, title, body, and CTAs."""

    def __init__(
        self,
        title: str,
        body: str = "",
        *,
        icon: str | None = "fa5s.inbox",
        primary_text: str | None = None,
        primary_callback: Callable[[], None] | None = None,
        secondary_text: str | None = None,
        secondary_callback: Callable[[], None] | None = None,
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("EmptyState")

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)
        outer.addStretch(1)

        if icon and _HAS_QTA:
            icon_label = QtWidgets.QLabel()
            icon_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            pix = _icon(icon).pixmap(48, 48)
            icon_label.setPixmap(pix)
            outer.addWidget(icon_label)

        title_label = QtWidgets.QLabel(title)
        title_label.setObjectName("EmptyStateTitle")
        title_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        title_label.setWordWrap(True)
        outer.addWidget(title_label)

        if body:
            body_label = QtWidgets.QLabel(body)
            body_label.setObjectName("EmptyStateBody")
            body_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
            body_label.setWordWrap(True)
            outer.addWidget(body_label)

        if primary_text or secondary_text:
            row = QtWidgets.QHBoxLayout()
            row.setSpacing(8)
            row.addStretch(1)
            if secondary_text:
                btn = QtWidgets.QPushButton(secondary_text)
                btn.setProperty("flat", "true")
                if secondary_callback:
                    btn.clicked.connect(secondary_callback)
                row.addWidget(btn)
            if primary_text:
                btn = QtWidgets.QPushButton(primary_text)
                btn.setProperty("primary", "true")
                btn.setDefault(True)
                if primary_callback:
                    btn.clicked.connect(primary_callback)
                row.addWidget(btn)
            row.addStretch(1)
            outer.addLayout(row)

        outer.addStretch(2)


class ErrorCard(QtWidgets.QFrame):
    """Inline error card with title, body, and a 'Copy details' button."""

    def __init__(
        self,
        title: str = "Something went wrong",
        body: str = "",
        *,
        details: str = "",
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("ErrorCard")
        self._details = details

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)

        header = QtWidgets.QHBoxLayout()
        header.setSpacing(8)
        if _HAS_QTA:
            icon_label = QtWidgets.QLabel()
            icon_label.setPixmap(_icon("fa5s.exclamation-triangle", color="#d73a49").pixmap(18, 18))
            header.addWidget(icon_label, 0, QtCore.Qt.AlignmentFlag.AlignTop)
        title_label = QtWidgets.QLabel(title)
        title_label.setProperty("role", "error-title")
        font = title_label.font()
        font.setBold(True)
        title_label.setFont(font)
        title_label.setWordWrap(True)
        header.addWidget(title_label, 1)
        layout.addLayout(header)

        if body:
            body_label = QtWidgets.QLabel(body)
            body_label.setWordWrap(True)
            layout.addWidget(body_label)

        if details:
            actions = QtWidgets.QHBoxLayout()
            actions.addStretch(1)
            copy_btn = QtWidgets.QPushButton("Copy details")
            copy_btn.setProperty("flat", "true")
            copy_btn.clicked.connect(self._copy_details)
            actions.addWidget(copy_btn)
            layout.addLayout(actions)

    def _copy_details(self) -> None:
        if self._details:
            QtWidgets.QApplication.clipboard().setText(self._details)


__all__ = ["EmptyState", "ErrorCard"]
