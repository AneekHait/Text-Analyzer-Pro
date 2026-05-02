"""Collapsible section widget with toggle header."""

from PySide6 import QtCore, QtWidgets


class CollapsibleSection(QtWidgets.QWidget):
    def __init__(self, title, subtitle="", expanded=True, parent=None):
        super().__init__(parent)
        self.toggle_button = QtWidgets.QToolButton(text=title)
        self.toggle_button.setCheckable(True)
        self.toggle_button.setChecked(expanded)
        self.toggle_button.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.toggle_button.setArrowType(QtCore.Qt.DownArrow if expanded else QtCore.Qt.RightArrow)
        self.subtitle_label = QtWidgets.QLabel(subtitle)
        self.subtitle_label.setWordWrap(True)
        self.content = QtWidgets.QWidget()
        self.content.setVisible(expanded)

        header_layout = QtWidgets.QVBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.addWidget(self.toggle_button)
        if subtitle:
            self.subtitle_label.setStyleSheet("color: #666666;")
            header_layout.addWidget(self.subtitle_label)

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addLayout(header_layout)
        outer.addWidget(self.content)

        self.toggle_button.toggled.connect(self._toggle)

    def _toggle(self, checked):
        self.toggle_button.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
        self.content.setVisible(checked)
