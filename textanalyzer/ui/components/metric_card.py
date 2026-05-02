"""MetricCard — compact stat display with label + value."""

from __future__ import annotations

from PySide6 import QtCore, QtWidgets

from ...utils.icons import _HAS_QTA, _icon


class MetricCard(QtWidgets.QFrame):
    """A small card showing one metric (icon, label, value, optional delta)."""

    def __init__(
        self,
        label: str,
        value: str = "—",
        *,
        icon: str | None = None,
        delta: str | None = None,
        parent: QtWidgets.QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("MetricCard")

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        header = QtWidgets.QHBoxLayout()
        header.setSpacing(6)

        if icon and _HAS_QTA:
            ic_label = QtWidgets.QLabel()
            ic_label.setPixmap(_icon(icon).pixmap(16, 16))
            header.addWidget(ic_label)

        self._label = QtWidgets.QLabel(label)
        self._label.setObjectName("MetricLabel")
        self._label.setProperty("role", "muted")
        header.addWidget(self._label, 1)
        layout.addLayout(header)

        self._value = QtWidgets.QLabel(value)
        self._value.setObjectName("MetricValue")
        val_font = self._value.font()
        val_font.setPointSize(val_font.pointSize() + 6)
        val_font.setBold(True)
        self._value.setFont(val_font)
        layout.addWidget(self._value)

        if delta:
            self._delta = QtWidgets.QLabel(delta)
            self._delta.setProperty("role", "muted")
            layout.addWidget(self._delta)
        else:
            self._delta = None

    def set_value(self, value: str) -> None:
        self._value.setText(value)

    def set_delta(self, text: str) -> None:
        if self._delta is not None:
            self._delta.setText(text)

    def value(self) -> str:
        return self._value.text()


__all__ = ["MetricCard"]
