"""Settings window — non-modal preferences dialog.

Categories:
- Appearance: theme mode, accent (placeholder).
- Defaults: clustering algorithm, top-N keywords, min cluster size.
- Data: recent file count, clear recent list.
- About: shortcut reference (read-only).

Persists via :mod:`app_settings`. Emits :pyattr:`settings_changed` whenever
a value is committed so the main window can react live (e.g., theme).
"""

from __future__ import annotations

from PySide6 import QtCore, QtGui, QtWidgets

import app_settings


class SettingsWindow(QtWidgets.QDialog):
    settings_changed = QtCore.Signal(dict)

    def __init__(self, settings: dict, parent: QtWidgets.QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.setObjectName("SettingsWindow")
        self.setModal(False)
        self.resize(620, 520)
        self._settings = settings

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(14)

        title = QtWidgets.QLabel("Settings")
        title.setProperty("role", "page-title")
        title_font = title.font()
        title_font.setPointSize(title_font.pointSize() + 4)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setObjectName("InnerTabs")
        layout.addWidget(self.tabs, 1)

        self.tabs.addTab(self._build_appearance_tab(), "Appearance")
        self.tabs.addTab(self._build_defaults_tab(), "Defaults")
        self.tabs.addTab(self._build_data_tab(), "Data")
        self.tabs.addTab(self._build_shortcuts_tab(), "Shortcuts")

        # Footer buttons.
        button_row = QtWidgets.QHBoxLayout()
        button_row.addStretch(1)
        self.close_btn = QtWidgets.QPushButton("Close")
        self.close_btn.setShortcut(QtGui.QKeySequence(QtCore.Qt.Key.Key_Escape))
        self.close_btn.clicked.connect(self.close)
        button_row.addWidget(self.close_btn)
        layout.addLayout(button_row)

    # ------------------------------------------------------------------
    # Tabs
    # ------------------------------------------------------------------
    def _build_appearance_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        form = QtWidgets.QFormLayout(page)
        form.setContentsMargins(16, 16, 16, 16)
        form.setSpacing(10)

        self.theme_combo = QtWidgets.QComboBox()
        self.theme_combo.addItems(["Dark", "Light"])
        current_theme = (self._settings.get("theme") or "dark").lower()
        self.theme_combo.setCurrentIndex(0 if current_theme == "dark" else 1)
        self.theme_combo.currentTextChanged.connect(self._on_theme_changed)
        form.addRow("Theme:", self.theme_combo)

        hint = QtWidgets.QLabel("Theme changes apply instantly to all open windows.")
        hint.setProperty("role", "muted")
        hint.setWordWrap(True)
        form.addRow(hint)

        return page

    def _build_defaults_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        form = QtWidgets.QFormLayout(page)
        form.setContentsMargins(16, 16, 16, 16)
        form.setSpacing(10)

        self.alg_combo = QtWidgets.QComboBox()
        self.alg_combo.addItems(["kmeans", "minibatchkmeans", "agglomerative", "dbscan"])
        last_alg = self._settings.get("last_algorithm") or "kmeans"
        idx = self.alg_combo.findText(last_alg)
        if idx >= 0:
            self.alg_combo.setCurrentIndex(idx)
        self.alg_combo.currentTextChanged.connect(self._on_default_alg_changed)
        form.addRow("Default algorithm:", self.alg_combo)

        self.topn_spin = QtWidgets.QSpinBox()
        self.topn_spin.setRange(3, 30)
        self.topn_spin.setValue(int(self._settings.get("default_top_n", 10) or 10))
        self.topn_spin.valueChanged.connect(self._on_topn_changed)
        form.addRow("Default top-N keywords:", self.topn_spin)

        return page

    def _build_data_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        recents = self._settings.get("recent_files") or []
        count_label = QtWidgets.QLabel(f"Recent files tracked: {len(recents)}")
        layout.addWidget(count_label)

        clear_btn = QtWidgets.QPushButton("Clear Recent Files")
        clear_btn.clicked.connect(self._clear_recent)
        layout.addWidget(clear_btn, 0, QtCore.Qt.AlignmentFlag.AlignLeft)
        self._recent_count_label = count_label

        layout.addStretch(1)
        return page

    def _build_shortcuts_tab(self) -> QtWidgets.QWidget:
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(8)

        rows = [
            ("Open file", "Ctrl+O"),
            ("Save results", "Ctrl+S"),
            ("Run clustering", "Ctrl+R"),
            ("Toggle theme", "Ctrl+T"),
            ("Open Settings", "Ctrl+,"),
            ("Open Diagnostics", "Ctrl+/"),
            ("Close tab", "Ctrl+W"),
            ("Next tab", "Ctrl+Tab"),
            ("Previous tab", "Ctrl+Shift+Tab"),
            ("About", "F1"),
            ("Quit", "Ctrl+Q"),
        ]
        table = QtWidgets.QTableWidget(len(rows), 2)
        table.setHorizontalHeaderLabels(["Action", "Shortcut"])
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(True)
        table.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
        table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.NoSelection)
        for i, (action, key) in enumerate(rows):
            table.setItem(i, 0, QtWidgets.QTableWidgetItem(action))
            table.setItem(i, 1, QtWidgets.QTableWidgetItem(key))
        table.resizeColumnsToContents()
        layout.addWidget(table, 1)
        return page

    # ------------------------------------------------------------------
    # Slots
    # ------------------------------------------------------------------
    def _on_theme_changed(self, value: str) -> None:
        mode = value.lower()
        self._settings["theme"] = mode
        app_settings.save(self._settings)
        self.settings_changed.emit({"theme": mode})

    def _on_default_alg_changed(self, value: str) -> None:
        self._settings["default_algorithm"] = value
        app_settings.save(self._settings)
        self.settings_changed.emit({"default_algorithm": value})

    def _on_topn_changed(self, value: int) -> None:
        self._settings["default_top_n"] = int(value)
        app_settings.save(self._settings)
        self.settings_changed.emit({"default_top_n": int(value)})

    def _clear_recent(self) -> None:
        self._settings["recent_files"] = []
        app_settings.save(self._settings)
        self._recent_count_label.setText("Recent files tracked: 0")
        self.settings_changed.emit({"recent_files": []})

    def sync_theme_from_external(self, mode: str) -> None:
        """Update the theme combo without re-emitting (used when toggled
        via Ctrl+T outside the dialog)."""
        self.theme_combo.blockSignals(True)
        self.theme_combo.setCurrentIndex(0 if mode == "dark" else 1)
        self.theme_combo.blockSignals(False)


__all__ = ["SettingsWindow"]
