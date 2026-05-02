"""Data source panel for the Setup page.

Two states stacked in a QStackedWidget:
  - DropZone: large drag-and-drop empty state shown when no file is loaded.
  - LoadedView: compact file summary, sheet selector, and a 5-row preview table.

Emits `file_picked(path)` when the user drops or browses a file. Sheet changes
emit `sheet_changed(sheet_name)`. The host wires these to the existing
`_open_file` / `_load_sheet` methods.
"""

from __future__ import annotations

import os
from typing import Iterable

import pandas as pd
from PySide6 import QtCore, QtGui, QtWidgets

from ..utils.icons import _icon


_SUPPORTED_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".ods", ".csv", ".json"}


def _format_size(num_bytes: float) -> str:
    if num_bytes < 1024:
        return f"{num_bytes:.0f} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / (1024 * 1024):.1f} MB"


class DropZone(QtWidgets.QFrame):
    """Big visual call-to-action: drop a file or click to browse."""

    file_picked = QtCore.Signal(str)
    browse_requested = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.setFixedHeight(120)

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(20, 14, 20, 14)
        layout.setSpacing(14)
        layout.setAlignment(QtCore.Qt.AlignCenter)

        icon_label = QtWidgets.QLabel()
        icon = _icon("fa5s.cloud-upload-alt")
        icon_label.setPixmap(icon.pixmap(32, 32))
        icon_label.setAlignment(QtCore.Qt.AlignCenter)
        layout.addWidget(icon_label)

        text_col = QtWidgets.QVBoxLayout()
        text_col.setSpacing(2)
        text_col.setAlignment(QtCore.Qt.AlignVCenter)

        title = QtWidgets.QLabel("Drop your data file here")
        title.setObjectName("DropZoneTitle")
        text_col.addWidget(title)

        sub = QtWidgets.QLabel("or click to browse · Excel, CSV, JSON, ODS")
        sub.setObjectName("DropZoneSubtitle")
        text_col.addWidget(sub)

        layout.addLayout(text_col)

    def sizeHint(self) -> QtCore.QSize:
        # setFixedHeight clamps the rendered size, but Qt's default sizeHint
        # ignores it — leaving parents (with Maximum policy) sized to the
        # layout's natural ~60 px and clipping the dropzone. Report the locked
        # height so the QGroupBox above us sizes correctly.
        return QtCore.QSize(super().sizeHint().width(), 120)

    def minimumSizeHint(self) -> QtCore.QSize:
        return QtCore.QSize(super().minimumSizeHint().width(), 120)

    def mousePressEvent(self, event: QtGui.QMouseEvent) -> None:
        if event.button() == QtCore.Qt.LeftButton:
            self.browse_requested.emit()
        super().mousePressEvent(event)

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:
        if self._first_supported_path(event.mimeData().urls()):
            event.acceptProposedAction()
            self.setProperty("dragging", "true")
            self._restyle()
        else:
            event.ignore()

    def dragLeaveEvent(self, event: QtCore.QEvent) -> None:
        self.setProperty("dragging", "false")
        self._restyle()
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QtGui.QDropEvent) -> None:
        path = self._first_supported_path(event.mimeData().urls())
        self.setProperty("dragging", "false")
        self._restyle()
        if path:
            event.acceptProposedAction()
            self.file_picked.emit(path)
        else:
            event.ignore()

    @staticmethod
    def _first_supported_path(urls: Iterable[QtCore.QUrl]) -> str | None:
        for url in urls:
            if not url.isLocalFile():
                continue
            local = url.toLocalFile()
            if os.path.splitext(local)[1].lower() in _SUPPORTED_EXTS:
                return local
        return None

    def _restyle(self) -> None:
        self.style().unpolish(self)
        self.style().polish(self)


class _CurrentPageStack(QtWidgets.QStackedWidget):
    """QStackedWidget that sizes to the *current* page, not the largest one."""

    def sizeHint(self) -> QtCore.QSize:
        widget = self.currentWidget()
        return widget.sizeHint() if widget is not None else super().sizeHint()

    def minimumSizeHint(self) -> QtCore.QSize:
        widget = self.currentWidget()
        return widget.minimumSizeHint() if widget is not None else super().minimumSizeHint()


class _PreviewModel(QtCore.QAbstractTableModel):
    def __init__(self, df: pd.DataFrame | None = None, parent=None) -> None:
        super().__init__(parent)
        self._df = df if df is not None else pd.DataFrame()
        self._highlight_col: str | None = None

    def set_dataframe(self, df: pd.DataFrame) -> None:
        self.beginResetModel()
        self._df = df
        self.endResetModel()

    def set_highlight_column(self, name: str | None) -> None:
        self._highlight_col = name
        if not self._df.empty:
            top = self.index(0, 0)
            bottom = self.index(self.rowCount() - 1, max(0, self.columnCount() - 1))
            self.dataChanged.emit(top, bottom, [QtCore.Qt.BackgroundRole])
        self.headerDataChanged.emit(QtCore.Qt.Horizontal, 0, max(0, self.columnCount() - 1))

    def rowCount(self, parent=QtCore.QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent=QtCore.QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df.columns)

    def data(self, index, role=QtCore.Qt.DisplayRole):
        if not index.isValid():
            return None
        if role == QtCore.Qt.DisplayRole:
            value = self._df.iat[index.row(), index.column()]
            if pd.isna(value):
                return ""
            text = str(value)
            return text if len(text) <= 120 else text[:117] + "…"
        if role == QtCore.Qt.ToolTipRole:
            value = self._df.iat[index.row(), index.column()]
            return "" if pd.isna(value) else str(value)
        return None

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        if role == QtCore.Qt.DisplayRole:
            if orientation == QtCore.Qt.Horizontal and 0 <= section < len(self._df.columns):
                return str(self._df.columns[section])
            if orientation == QtCore.Qt.Vertical:
                return str(section + 1)
        if role == QtCore.Qt.FontRole and orientation == QtCore.Qt.Horizontal:
            if self._highlight_col and 0 <= section < len(self._df.columns):
                if str(self._df.columns[section]) == self._highlight_col:
                    font = QtGui.QFont()
                    font.setBold(True)
                    return font
        return None


class DataSourcePanel(QtWidgets.QFrame):
    """Stacked panel: dropzone empty state ↔ loaded-file view with preview."""

    file_picked = QtCore.Signal(str)
    browse_requested = QtCore.Signal()
    sheet_changed = QtCore.Signal(str)
    change_requested = QtCore.Signal()

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("DataSourcePanel")
        self.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Maximum)

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._stack = _CurrentPageStack()
        self._stack.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Maximum)
        self._stack.currentChanged.connect(lambda *_: self.updateGeometry())
        outer.addWidget(self._stack)

        self.dropzone = DropZone()
        self.dropzone.file_picked.connect(self.file_picked.emit)
        self.dropzone.browse_requested.connect(self.browse_requested.emit)
        self._stack.addWidget(self.dropzone)

        self._loaded = self._build_loaded_view()
        self._stack.addWidget(self._loaded)

    # ------------------------------------------------------------------ #
    # Public API                                                         #
    # ------------------------------------------------------------------ #

    def show_empty_state(self) -> None:
        self._stack.setCurrentWidget(self.dropzone)

    def show_loaded_state(self, *, file_path: str, sheets: list[str], current_sheet: str) -> None:
        self._file_name_label.setText(os.path.basename(file_path))
        try:
            size = os.path.getsize(file_path)
            self._file_meta_label.setText(_format_size(size))
        except OSError:
            self._file_meta_label.setText("")
        self._sheet_combo.blockSignals(True)
        self._sheet_combo.clear()
        self._sheet_combo.addItems(sheets)
        if current_sheet in sheets:
            self._sheet_combo.setCurrentText(current_sheet)
        self._sheet_combo.blockSignals(False)
        self._sheet_combo.setVisible(len(sheets) > 1)
        self._sheet_label.setVisible(len(sheets) > 1)
        self._stack.setCurrentWidget(self._loaded)

    def update_preview(self, df: pd.DataFrame, *, highlight_column: str | None = None) -> None:
        sample = df.head(5).copy()
        self._preview_model.set_dataframe(sample)
        self._preview_model.set_highlight_column(highlight_column)
        rows, cols = len(df), len(df.columns)
        self._shape_label.setText(f"{rows:,} rows · {cols} columns")
        self._preview_table.resizeColumnsToContents()
        for i in range(self._preview_model.columnCount()):
            current = self._preview_table.columnWidth(i)
            self._preview_table.setColumnWidth(i, min(current, 220))

    def set_highlight_column(self, name: str | None) -> None:
        self._preview_model.set_highlight_column(name)

    def current_sheet(self) -> str:
        return self._sheet_combo.currentText()

    # ------------------------------------------------------------------ #
    # Internal                                                           #
    # ------------------------------------------------------------------ #

    def _build_loaded_view(self) -> QtWidgets.QWidget:
        wrapper = QtWidgets.QWidget()
        # Summary row ~46 + spacing 10 + preview table min 150 = 206. Pin so
        # parents with Maximum size policy don't clip the table.
        wrapper.setMinimumHeight(210)
        layout = QtWidgets.QVBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # Summary row
        summary = QtWidgets.QFrame()
        summary.setObjectName("DataSourceSummary")
        summary_layout = QtWidgets.QHBoxLayout(summary)
        summary_layout.setContentsMargins(14, 10, 14, 10)
        summary_layout.setSpacing(10)

        icon_label = QtWidgets.QLabel()
        icon_label.setPixmap(_icon("fa5s.file-alt").pixmap(20, 20))
        summary_layout.addWidget(icon_label)

        self._file_name_label = QtWidgets.QLabel("")
        self._file_name_label.setObjectName("DataSourceFile")
        summary_layout.addWidget(self._file_name_label)

        self._file_meta_label = QtWidgets.QLabel("")
        self._file_meta_label.setProperty("role", "muted")
        summary_layout.addWidget(self._file_meta_label)

        sep = QtWidgets.QLabel("·")
        sep.setProperty("role", "muted")
        summary_layout.addWidget(sep)

        self._shape_label = QtWidgets.QLabel("")
        self._shape_label.setProperty("role", "muted")
        summary_layout.addWidget(self._shape_label)

        summary_layout.addStretch(1)

        self._sheet_label = QtWidgets.QLabel("Sheet:")
        self._sheet_label.setProperty("role", "muted")
        summary_layout.addWidget(self._sheet_label)

        self._sheet_combo = QtWidgets.QComboBox()
        self._sheet_combo.setMinimumWidth(160)
        self._sheet_combo.currentTextChanged.connect(self.sheet_changed.emit)
        summary_layout.addWidget(self._sheet_combo)

        change_btn = QtWidgets.QPushButton("Change")
        change_btn.setProperty("flat", "true")
        change_btn.setIcon(_icon("fa5s.exchange-alt"))
        change_btn.clicked.connect(self.change_requested.emit)
        summary_layout.addWidget(change_btn)

        layout.addWidget(summary)

        # Preview table
        self._preview_model = _PreviewModel()
        self._preview_table = QtWidgets.QTableView()
        self._preview_table.setObjectName("PreviewTable")
        self._preview_table.setModel(self._preview_model)
        self._preview_table.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self._preview_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self._preview_table.setFocusPolicy(QtCore.Qt.NoFocus)
        self._preview_table.setAlternatingRowColors(True)
        self._preview_table.verticalHeader().setDefaultSectionSize(24)
        self._preview_table.horizontalHeader().setStretchLastSection(True)
        self._preview_table.setMinimumHeight(150)
        self._preview_table.setMaximumHeight(180)
        layout.addWidget(self._preview_table)

        return wrapper
