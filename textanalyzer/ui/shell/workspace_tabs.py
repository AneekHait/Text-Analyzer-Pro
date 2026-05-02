"""Workspace tab bar — the top-level multi-document container.

Currently used in single-tab mode (Phase 2). Phase 3 will add real
multi-session support; this class is already prepared for it: tabs are
closeable, renamable (double-click title), and reorderable, with a
right-click context menu for batch actions.
"""

from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets


class WorkspaceTabWidget(QtWidgets.QTabWidget):
    """QTabWidget styled as the workspace's document tab bar."""

    tab_close_requested = QtCore.Signal(int)
    tab_duplicate_requested = QtCore.Signal(int)
    tab_close_others_requested = QtCore.Signal(int)
    tab_close_right_requested = QtCore.Signal(int)
    tab_rename_requested = QtCore.Signal(int, str)

    def __init__(self, parent: QtWidgets.QWidget | None = None):
        super().__init__(parent)
        self.setObjectName("WorkspaceTabs")
        self.setDocumentMode(True)
        self.setMovable(True)
        self.setTabsClosable(True)
        self.setUsesScrollButtons(True)
        self.setElideMode(QtCore.Qt.TextElideMode.ElideRight)

        self.tabBar().setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.tabBar().customContextMenuRequested.connect(self._on_tab_context_menu)
        self.tabBar().tabBarDoubleClicked.connect(self._on_tab_double_clicked)
        self.tabCloseRequested.connect(self.tab_close_requested.emit)
        self._update_tab_bar_visibility()

    # ------------------------------------------------------------------
    # Tab bar visibility — hide the strip when there's only one document.
    # When Phase 3 lands and a second tab opens, the bar reappears
    # automatically. No call sites need to change.
    # ------------------------------------------------------------------
    def tabInserted(self, index: int) -> None:  # noqa: N802 (Qt override)
        super().tabInserted(index)
        self._update_tab_bar_visibility()

    def tabRemoved(self, index: int) -> None:  # noqa: N802 (Qt override)
        super().tabRemoved(index)
        self._update_tab_bar_visibility()

    def _update_tab_bar_visibility(self) -> None:
        self.tabBar().setVisible(self.count() > 1)

    # ------------------------------------------------------------------
    # Tab creation helpers
    # ------------------------------------------------------------------
    def add_workspace_tab(self, widget: QtWidgets.QWidget, title: str) -> int:
        """Append a tab and return its index."""
        index = self.addTab(widget, title)
        self.setCurrentIndex(index)
        return index

    # ------------------------------------------------------------------
    # Context menu / rename
    # ------------------------------------------------------------------
    def _on_tab_context_menu(self, pos: QtCore.QPoint) -> None:
        index = self.tabBar().tabAt(pos)
        if index < 0:
            return
        menu = QtWidgets.QMenu(self)
        close = menu.addAction("Close")
        close_others = menu.addAction("Close Others")
        close_right = menu.addAction("Close to the Right")
        menu.addSeparator()
        duplicate = menu.addAction("Duplicate")
        rename = menu.addAction("Rename…")
        chosen = menu.exec(self.tabBar().mapToGlobal(pos))
        if chosen is None:
            return
        if chosen is close:
            self.tab_close_requested.emit(index)
        elif chosen is close_others:
            self.tab_close_others_requested.emit(index)
        elif chosen is close_right:
            self.tab_close_right_requested.emit(index)
        elif chosen is duplicate:
            self.tab_duplicate_requested.emit(index)
        elif chosen is rename:
            self._prompt_rename(index)

    def _on_tab_double_clicked(self, index: int) -> None:
        if index >= 0:
            self._prompt_rename(index)

    def _prompt_rename(self, index: int) -> None:
        if index < 0 or index >= self.count():
            return
        current = self.tabText(index)
        new_name, ok = QtWidgets.QInputDialog.getText(
            self, "Rename Tab", "Tab name:", QtWidgets.QLineEdit.EchoMode.Normal, current
        )
        if ok:
            cleaned = new_name.strip()
            if cleaned and cleaned != current:
                self.setTabText(index, cleaned)
                self.tab_rename_requested.emit(index, cleaned)


__all__ = ["WorkspaceTabWidget"]
