"""Lightweight non-blocking toast notifications.

A :class:`ToastManager` (per top-level window) queues toasts and stacks
up to ``max_visible`` of them in the bottom-right corner. Older toasts
slide up as new ones appear; surplus toasts wait in a FIFO queue.
"""

from PySide6 import QtCore, QtWidgets

from ..utils.icons import _HAS_QTA, _icon


class Toast(QtWidgets.QFrame):
    """Lightweight non-blocking notification popping in the bottom-right."""

    closed = QtCore.Signal(object)

    def __init__(self, parent, message: str, level: str = "info", timeout: int = 3200):
        super().__init__(parent)
        self.setObjectName("Toast")
        self.setProperty("level", level)
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.setWindowFlags(QtCore.Qt.SubWindow | QtCore.Qt.FramelessWindowHint)

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(10)

        icon_map = {
            "success": ("fa5s.check-circle", "#1f9d55"),
            "warning": ("fa5s.exclamation-triangle", "#d29922"),
            "danger": ("fa5s.times-circle", "#d73a49"),
            "info": ("fa5s.info-circle", "#3d6dff"),
        }
        icon_name, color = icon_map.get(level, icon_map["info"])
        if _HAS_QTA:
            icon_label = QtWidgets.QLabel()
            icon_label.setPixmap(_icon(icon_name, color=color).pixmap(18, 18))
            layout.addWidget(icon_label)

        text_label = QtWidgets.QLabel(message)
        text_label.setWordWrap(True)
        layout.addWidget(text_label, 1)

        self._timeout_timer = QtCore.QTimer(self)
        self._timeout_timer.setSingleShot(True)
        self._timeout_timer.timeout.connect(self.close)
        self._timeout_timer.start(timeout)

    def closeEvent(self, event):  # type: ignore[override]
        try:
            self.closed.emit(self)
        finally:
            super().closeEvent(event)

    def show_at(self, anchor_widget: QtWidgets.QWidget, *, offset_y: int = 0):
        self.adjustSize()
        if anchor_widget is None:
            return
        geo = anchor_widget.geometry()
        margin = 18
        x = geo.right() - self.width() - margin
        y = geo.bottom() - self.height() - margin - 28 - offset_y
        self.move(x, y)
        self.show()
        self.raise_()


class ToastManager(QtCore.QObject):
    """Stacks up to ``max_visible`` toasts on a target window."""

    _attr = "_tap_toast_manager"

    def __init__(self, target: QtWidgets.QWidget, *, max_visible: int = 3, gap: int = 8):
        super().__init__(target)
        self.target = target
        self.max_visible = max_visible
        self.gap = gap
        self._visible: list[Toast] = []
        self._queue: list[tuple[str, str, int]] = []

    @classmethod
    def for_window(cls, widget: QtWidgets.QWidget) -> "ToastManager | None":
        window = widget.window() if widget is not None else None
        target = window if window is not None else widget
        if target is None:
            return None
        existing = getattr(target, cls._attr, None)
        if isinstance(existing, ToastManager):
            return existing
        manager = cls(target)
        setattr(target, cls._attr, manager)
        return manager

    def show(self, message: str, level: str = "info", timeout: int = 3200) -> None:
        if len(self._visible) >= self.max_visible:
            self._queue.append((message, level, timeout))
            return
        toast = Toast(self.target, message, level=level, timeout=timeout)
        toast.closed.connect(self._on_closed)
        self._visible.append(toast)
        self._reposition()

    def dismiss_all(self) -> None:
        self._queue.clear()
        for toast in list(self._visible):
            toast.close()

    def _on_closed(self, toast: Toast) -> None:
        if toast in self._visible:
            self._visible.remove(toast)
        self._reposition()
        # Pop next from queue if any.
        if self._queue and len(self._visible) < self.max_visible:
            message, level, timeout = self._queue.pop(0)
            self.show(message, level=level, timeout=timeout)

    def _reposition(self) -> None:
        offset = 0
        # Stack from bottom up.
        for toast in self._visible:
            toast.show_at(self.target, offset_y=offset)
            offset += toast.height() + self.gap


def show_toast(parent: QtWidgets.QWidget, message: str, level: str = "info"):
    """Show a toast on the top-level window of *parent* via the manager."""
    manager = ToastManager.for_window(parent)
    if manager is None:
        return
    manager.show(message, level=level)

