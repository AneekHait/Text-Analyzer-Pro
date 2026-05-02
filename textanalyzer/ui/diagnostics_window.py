"""Diagnostics window — non-modal log viewer.

Mirrors entries appended via :meth:`ClusterGUI.log_msg` into a dedicated
window with filter, search, copy, and clear actions. The main window's
in-page log remains as the source of truth; this window is a viewer that
can outlive page navigation.
"""

from __future__ import annotations

from datetime import datetime

from PySide6 import QtCore, QtGui, QtWidgets


class DiagnosticsWindow(QtWidgets.QDialog):
    cleared = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Diagnostics")
        self.setObjectName("DiagnosticsWindow")
        self.setModal(False)
        self.resize(720, 500)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 12)
        layout.setSpacing(10)

        title = QtWidgets.QLabel("Diagnostics")
        title_font = title.font()
        title_font.setPointSize(title_font.pointSize() + 3)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        # Toolbar row.
        toolbar = QtWidgets.QHBoxLayout()
        toolbar.setSpacing(8)

        self.filter_edit = QtWidgets.QLineEdit()
        self.filter_edit.setPlaceholderText("Filter log entries…")
        self.filter_edit.setClearButtonEnabled(True)
        self.filter_edit.textChanged.connect(self._refilter)
        toolbar.addWidget(self.filter_edit, 1)

        copy_btn = QtWidgets.QPushButton("Copy All")
        copy_btn.clicked.connect(self._copy_all)
        toolbar.addWidget(copy_btn)

        clear_btn = QtWidgets.QPushButton("Clear")
        clear_btn.clicked.connect(self._clear)
        toolbar.addWidget(clear_btn)

        layout.addLayout(toolbar)

        self.view = QtWidgets.QPlainTextEdit()
        self.view.setObjectName("DiagnosticsView")
        self.view.setReadOnly(True)
        self.view.setLineWrapMode(QtWidgets.QPlainTextEdit.LineWrapMode.NoWrap)
        layout.addWidget(self.view, 1)

        # Footer.
        footer = QtWidgets.QHBoxLayout()
        footer.addStretch(1)
        close_btn = QtWidgets.QPushButton("Close")
        close_btn.setShortcut(QtGui.QKeySequence(QtCore.Qt.Key.Key_Escape))
        close_btn.clicked.connect(self.close)
        footer.addWidget(close_btn)
        layout.addLayout(footer)

        self._all_lines: list[str] = []

    # ------------------------------------------------------------------
    # Public API used by the main window
    # ------------------------------------------------------------------
    def append_entry(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}"
        self._all_lines.append(line)
        if self._matches_filter(line):
            self.view.appendPlainText(line)

    def set_history(self, lines: list[str]) -> None:
        """Replace contents with an existing log history (one line per entry)."""
        self._all_lines = list(lines)
        self._rebuild_view()

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------
    def _matches_filter(self, line: str) -> bool:
        needle = self.filter_edit.text().strip().lower()
        if not needle:
            return True
        return needle in line.lower()

    def _refilter(self) -> None:
        self._rebuild_view()

    def _rebuild_view(self) -> None:
        self.view.blockSignals(True)
        self.view.clear()
        if self._all_lines:
            visible = [ln for ln in self._all_lines if self._matches_filter(ln)]
            self.view.setPlainText("\n".join(visible))
        self.view.blockSignals(False)

    def _copy_all(self) -> None:
        text = self.view.toPlainText()
        QtWidgets.QApplication.clipboard().setText(text)

    def _clear(self) -> None:
        self._all_lines.clear()
        self.view.clear()
        self.cleared.emit()


__all__ = ["DiagnosticsWindow"]
