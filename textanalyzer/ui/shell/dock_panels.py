"""Dock panels for the workspace shell.

* :class:`NavigatorDock` (left) — recent files, open sessions, quick links.
* :class:`InspectorDock` (right) — context-sensitive details for the
  currently selected item (column, cluster row, …).

Both are designed to be optional; the analysis tab remains fully usable
when the docks are hidden.
"""

from __future__ import annotations

from typing import Callable

from PySide6 import QtCore, QtGui, QtWidgets


class _BaseDock(QtWidgets.QDockWidget):
    """Shared styling for shell docks."""

    def __init__(self, title: str, object_name: str, parent: QtWidgets.QWidget | None = None):
        super().__init__(title, parent)
        self.setObjectName(object_name)
        self.setFeatures(
            QtWidgets.QDockWidget.DockWidgetFeature.DockWidgetClosable
            | QtWidgets.QDockWidget.DockWidgetFeature.DockWidgetMovable
            | QtWidgets.QDockWidget.DockWidgetFeature.DockWidgetFloatable
        )
        self.setAllowedAreas(
            QtCore.Qt.DockWidgetArea.LeftDockWidgetArea
            | QtCore.Qt.DockWidgetArea.RightDockWidgetArea
        )


class NavigatorDock(_BaseDock):
    """Left-side navigator: recent files & open sessions."""

    file_activated = QtCore.Signal(str)

    def __init__(self, parent: QtWidgets.QWidget | None = None):
        super().__init__("Navigator", "NavigatorDock", parent)
        body = QtWidgets.QWidget()
        body.setObjectName("NavigatorBody")
        layout = QtWidgets.QVBoxLayout(body)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        recent_label = QtWidgets.QLabel("Recent Files")
        recent_label.setProperty("role", "section-title")
        layout.addWidget(recent_label)

        self.recent_list = QtWidgets.QListWidget()
        self.recent_list.setObjectName("RecentList")
        self.recent_list.setAlternatingRowColors(False)
        self.recent_list.setUniformItemSizes(True)
        self.recent_list.itemActivated.connect(self._on_item_activated)
        self.recent_list.itemDoubleClicked.connect(self._on_item_activated)
        layout.addWidget(self.recent_list, 1)

        empty_hint = QtWidgets.QLabel("No recent files yet.\nOpen a CSV to get started.")
        empty_hint.setObjectName("RecentEmptyHint")
        empty_hint.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        empty_hint.setWordWrap(True)
        empty_hint.setProperty("role", "muted")
        layout.addWidget(empty_hint)
        self._empty_hint = empty_hint
        self._update_empty_state()

        self.setWidget(body)

    def set_recent_files(self, files: list[str]) -> None:
        self.recent_list.clear()
        import os
        for path in files:
            item = QtWidgets.QListWidgetItem(os.path.basename(path) or path)
            item.setToolTip(path)
            item.setData(QtCore.Qt.ItemDataRole.UserRole, path)
            self.recent_list.addItem(item)
        self._update_empty_state()

    def _update_empty_state(self) -> None:
        empty = self.recent_list.count() == 0
        self._empty_hint.setVisible(empty)
        self.recent_list.setVisible(not empty)

    def _on_item_activated(self, item: QtWidgets.QListWidgetItem) -> None:
        path = item.data(QtCore.Qt.ItemDataRole.UserRole)
        if path:
            self.file_activated.emit(str(path))


class InspectorDock(_BaseDock):
    """Right-side inspector showing details for the current selection."""

    def __init__(self, parent: QtWidgets.QWidget | None = None):
        super().__init__("Inspector", "InspectorDock", parent)
        body = QtWidgets.QWidget()
        body.setObjectName("InspectorBody")
        layout = QtWidgets.QVBoxLayout(body)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)

        self.title = QtWidgets.QLabel("Inspector")
        self.title.setProperty("role", "section-title")
        layout.addWidget(self.title)

        self.body_label = QtWidgets.QLabel(
            "Select a column, cluster, or row to see details here."
        )
        self.body_label.setWordWrap(True)
        self.body_label.setAlignment(
            QtCore.Qt.AlignmentFlag.AlignTop | QtCore.Qt.AlignmentFlag.AlignLeft
        )
        self.body_label.setProperty("role", "muted")
        layout.addWidget(self.body_label, 1)

        self.setWidget(body)

    def set_details(self, title: str, body: str) -> None:
        self.title.setText(title or "Inspector")
        self.body_label.setText(body or "")


__all__ = ["NavigatorDock", "InspectorDock"]
