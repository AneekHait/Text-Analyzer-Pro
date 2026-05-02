"""ThemeManager: applies QSS to QApplication and notifies listeners on change."""
from PySide6 import QtCore, QtGui, QtWidgets

from .qss import build_qss
from .tokens import get_tokens


_FONT_FAMILIES_CACHE: "frozenset[str] | None" = None


def _system_font_families() -> "frozenset[str]":
    """Return the system's installed font family names as a cached frozenset.

    ``QFontDatabase().families()`` allocates a fresh list and on Windows
    can take ~30-50 ms the first time it touches the registry. Cache the
    result for the process lifetime — fonts don't appear/disappear at
    runtime in any scenario this app cares about.
    """
    global _FONT_FAMILIES_CACHE
    if _FONT_FAMILIES_CACHE is None:
        _FONT_FAMILIES_CACHE = frozenset(QtGui.QFontDatabase().families())
    return _FONT_FAMILIES_CACHE


class ThemeManager(QtCore.QObject):
    """Singleton-like helper for applying and toggling app themes."""

    changed = QtCore.Signal(str)  # emits new mode name

    def __init__(self, app: QtWidgets.QApplication, mode: str = "light"):
        super().__init__(app)
        self._app = app
        self._mode = "dark" if str(mode).lower() == "dark" else "light"

    @property
    def mode(self) -> str:
        return self._mode

    def tokens(self) -> dict:
        return get_tokens(self._mode)

    def apply(self, mode: str = None) -> None:
        if mode is not None:
            self._mode = "dark" if str(mode).lower() == "dark" else "light"
        self._app.setStyleSheet(build_qss(self._mode))
        self._apply_palette()
        self._apply_font()
        self.changed.emit(self._mode)

    def toggle(self) -> str:
        self.apply("light" if self._mode == "dark" else "dark")
        return self._mode

    def _apply_font(self) -> None:
        """Set the application font based on platform-aware tokens.

        Iterates the CSS-style font stack and picks the first family
        actually installed on the system so Linux / macOS / Windows all
        get the best available typeface.
        """
        t = self.tokens()
        family_stack = t.get("font_family", "")

        # Parse "'Family A', 'Family B', sans-serif" into a list.
        candidates: list[str] = []
        for part in family_stack.split(","):
            name = part.strip().strip("'\"")
            if name and name not in ("sans-serif", "system-ui", "monospace"):
                candidates.append(name)

        # Resolve the first installed family. Use the cached frozenset for
        # O(1) lookup instead of a list scan inside the loop, and skip the
        # QFontDatabase round-trip entirely on subsequent theme toggles.
        families = _system_font_families()
        chosen = next((n for n in candidates if n in families), "")
        if not chosen:
            chosen = candidates[0] if candidates else "Segoe UI"

        size_pt = int(t.get("_base_font_pt", 13))
        font = QtGui.QFont(chosen)
        font.setPixelSize(size_pt)
        self._app.setFont(font)

    def _apply_palette(self) -> None:
        t = self.tokens()
        palette = QtGui.QPalette()
        palette.setColor(QtGui.QPalette.Window, QtGui.QColor(t["bg"]))
        palette.setColor(QtGui.QPalette.WindowText, QtGui.QColor(t["text"]))
        palette.setColor(QtGui.QPalette.Base, QtGui.QColor(t["input_bg"]))
        palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(t["surface_alt"]))
        palette.setColor(QtGui.QPalette.Text, QtGui.QColor(t["text"]))
        palette.setColor(QtGui.QPalette.Button, QtGui.QColor(t["surface"]))
        palette.setColor(QtGui.QPalette.ButtonText, QtGui.QColor(t["text"]))
        palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(t["selection_bg"]))
        palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(t["selection_fg"]))
        palette.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor(t["tooltip_bg"]))
        palette.setColor(QtGui.QPalette.ToolTipText, QtGui.QColor(t["tooltip_fg"]))
        palette.setColor(QtGui.QPalette.PlaceholderText, QtGui.QColor(t["placeholder"]))
        self._app.setPalette(palette)

    @staticmethod
    def detect_system_mode() -> str:
        """Best-effort OS color-scheme detection. Returns 'light' or 'dark'."""
        try:
            hints = QtGui.QGuiApplication.styleHints()
            scheme = getattr(hints, "colorScheme", None)
            if scheme is not None:
                value = scheme()
                if int(value) == int(QtCore.Qt.ColorScheme.Dark):
                    return "dark"
        except Exception:
            pass
        return "light"
