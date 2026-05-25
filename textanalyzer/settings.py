"""User settings persisted to ~/.text_analyzer_pro/settings.json.

A lightweight QSettings mirror keeps OS-level integrations (window-restore,
macOS App Sandbox, Windows registry) in sync for the keys that benefit from it.
"""
import json
import os
from pathlib import Path

SETTINGS_DIR = Path.home() / ".text_analyzer_pro"
SETTINGS_FILE = SETTINGS_DIR / "settings.json"

DEFAULTS = {
    "theme": "system",  # 'light' | 'dark' | 'system'
    "geometry_main": None,  # base64 string from QByteArray
    "geometry_wordcloud": None,
    "window_state": None,  # base64 string from saveState()
    "recent_files": [],  # list[str]
    "max_recent": 8,
    "last_column": "",
    "last_algorithm": "kmeans",
    "cleaning_recipes": {},  # name -> dict
    "last_vectorizer_kind": "tfidf",  # 'tfidf' | 'embedding'
    "embedding_model": "sentence-transformers/all-MiniLM-L6-v2",
    "embedding_device": "cpu",  # 'cpu' | 'cuda'
    "embedding_batch_size": 32,
    "last_min_cluster_size": 5,
    "last_non_repetitive_min_size": 5,
    "last_granularity": 50,
    "taxonomy_confidence_threshold": 0.45,
    "user_taxonomy_renames": {},  # fingerprint -> edited subcategory name
}

# Keys mirrored to QSettings for OS-level integration (geometry restore,
# registry on Windows, plist on macOS).
_QSETTINGS_KEYS = ("theme", "geometry_main", "window_state")


def _load_raw() -> dict:
    if not SETTINGS_FILE.exists():
        return {}
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            return data
    except Exception:
        return {}
    return {}


def load() -> dict:
    raw = _load_raw()
    # Fall back to QSettings for mirrored keys when JSON is absent.
    if not raw:
        raw = _load_from_qsettings()
    merged = dict(DEFAULTS)
    merged.update({k: v for k, v in raw.items() if k in DEFAULTS})
    # ensure expected types
    if not isinstance(merged.get("recent_files"), list):
        merged["recent_files"] = []
    if not isinstance(merged.get("cleaning_recipes"), dict):
        merged["cleaning_recipes"] = {}
    if not isinstance(merged.get("user_taxonomy_renames"), dict):
        merged["user_taxonomy_renames"] = {}
    return merged


def save(data: dict) -> None:
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        clean = {k: data.get(k, v) for k, v in DEFAULTS.items()}
        with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
            json.dump(clean, fh, indent=2, ensure_ascii=False)
    except Exception:
        pass
    _mirror_to_qsettings(data)


def add_recent_file(data: dict, path: str) -> None:
    if not path:
        return
    path = os.path.abspath(path)
    recents = [p for p in data.get("recent_files", []) if p != path]
    recents.insert(0, path)
    max_recent = int(data.get("max_recent", DEFAULTS["max_recent"]))
    data["recent_files"] = recents[:max_recent]


# ------------------------------------------------------------------ #
# QSettings mirror                                                   #
# ------------------------------------------------------------------ #

def _qsettings():
    """Return a QSettings instance scoped to the application.

    Only call after QApplication has been created.
    """
    try:
        from PySide6.QtCore import QSettings
        return QSettings("Aneek Hait", "Text Analyzer Pro")
    except Exception:
        return None


def _mirror_to_qsettings(data: dict) -> None:
    """Write selected keys to the platform's native settings store."""
    qs = _qsettings()
    if qs is None:
        return
    for key in _QSETTINGS_KEYS:
        val = data.get(key)
        if val is not None:
            qs.setValue(key, val)
        else:
            qs.remove(key)
    qs.sync()


def _load_from_qsettings() -> dict:
    """Read mirrored keys back from QSettings (fallback source)."""
    qs = _qsettings()
    if qs is None:
        return {}
    out = {}
    for key in _QSETTINGS_KEYS:
        val = qs.value(key)
        if val is not None:
            out[key] = val
    return out
