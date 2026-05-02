"""Icon and asset helpers (qtawesome wrappers + branded banner cache)."""

from pathlib import Path

from PySide6 import QtCore, QtGui

try:
    import qtawesome as qta  # type: ignore
    _HAS_QTA = True
except Exception:  # pragma: no cover
    qta = None
    _HAS_QTA = False


ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets"


def _icon(name: str, color: str = None) -> QtGui.QIcon:
    """Return a qtawesome icon, or an empty icon when qtawesome is unavailable."""
    if not _HAS_QTA:
        return QtGui.QIcon()
    try:
        if color:
            return qta.icon(name, color=color)
        return qta.icon(name)
    except Exception:
        return QtGui.QIcon()


_ROLE_TOKEN_MAP = {
    "default": "text",
    "muted": "text_muted",
    "subtle": "text_subtle",
    "accent": "accent",
    "danger": "danger",
    "success": "success",
    "warning": "warning",
    "inverse": "text_inverse",
    "sidebar": "sidebar_fg",
    "sidebar_active": "sidebar_fg_active",
}


def themed_icon(name: str, role: str = "default", mode: str = "light") -> QtGui.QIcon:
    """Return a qtawesome icon coloured from theme tokens.

    ``role`` selects which token drives the icon colour. Falls back to
    a plain icon if qtawesome or token lookup fails.
    """
    if not _HAS_QTA:
        return QtGui.QIcon()
    try:
        from theme.tokens import get_tokens

        token_key = _ROLE_TOKEN_MAP.get(role, "text")
        color = get_tokens(mode).get(token_key, None)
        if color:
            return qta.icon(name, color=color)
    except Exception:
        pass
    return _icon(name)


def _spin_icon(name: str, color: str = None):
    """Return an animated qtawesome icon spinner or None if unavailable."""
    if not _HAS_QTA:
        return None
    try:
        kwargs = {"animation": qta.Spin}
        if color:
            kwargs["color"] = color
        return qta.icon(name, **kwargs)
    except Exception:
        return None


def _load_app_icon() -> QtGui.QIcon:
    """Load app icon from assets/, supporting ico/png/svg fallbacks."""
    icon = QtGui.QIcon()
    for fname in ("logo.ico", "logo.png", "logo_256.png", "logo.svg"):
        path = ASSETS_DIR / fname
        if path.exists():
            icon.addFile(str(path))
    return icon


_BANNER_CACHE: "dict[str, QtGui.QPixmap]" = {}


def _load_banner_pixmap(target_height: int) -> QtGui.QPixmap:
    """Load the trimmed brand banner scaled to the requested height."""
    cache_key = f"banner@{target_height}"
    if cache_key in _BANNER_CACHE:
        return _BANNER_CACHE[cache_key]
    for fname in ("logo_banner_trim.png", "logo_banner.png"):
        path = ASSETS_DIR / fname
        if path.exists():
            pixmap = QtGui.QPixmap(str(path))
            if not pixmap.isNull():
                pixmap = pixmap.scaledToHeight(target_height, QtCore.Qt.SmoothTransformation)
                _BANNER_CACHE[cache_key] = pixmap
                return pixmap
    return QtGui.QPixmap()
