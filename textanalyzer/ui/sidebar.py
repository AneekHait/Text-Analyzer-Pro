"""Sidebar navigation button used by the main shell."""

from PySide6 import QtCore, QtWidgets

from ..utils.icons import _HAS_QTA, _icon


class SidebarButton(QtWidgets.QToolButton):
    def __init__(self, text: str, icon_name: str = "", parent=None):
        super().__init__(parent)
        self.setObjectName("SidebarNav")
        self.setText("  " + text)
        self.setCheckable(True)
        self.setAutoExclusive(False)
        self.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        if icon_name and _HAS_QTA:
            self.setIcon(_icon(icon_name, color="#cfd3dc"))
            self.setIconSize(QtCore.QSize(18, 18))
        self.setMinimumHeight(38)
