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

from textanalyzer import settings as app_settings


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

        # Default vectorizer kind. Embedding mode is greyed-out at runtime if
        # sentence-transformers isn't importable — mirrors the Setup-tab combo.
        try:
            from textanalyzer.engine.cluster import _ST_AVAILABLE as _ST_OK
        except Exception:
            _ST_OK = False
        self.vec_combo = QtWidgets.QComboBox()
        self.vec_combo.addItem("TF-IDF (lexical)", userData="tfidf")
        self.vec_combo.addItem("Embeddings (semantic)", userData="embedding")
        if not _ST_OK:
            model = self.vec_combo.model()
            item = model.item(1) if hasattr(model, "item") else None
            if item is not None:
                item.setEnabled(False)
            self.vec_combo.setItemData(
                1,
                "Requires `pip install sentence-transformers`.",
                QtCore.Qt.ItemDataRole.ToolTipRole,
            )
        last_vec = (self._settings.get("last_vectorizer_kind") or "tfidf").lower()
        if last_vec == "embedding" and _ST_OK:
            self.vec_combo.setCurrentIndex(1)
        self.vec_combo.currentIndexChanged.connect(self._on_default_vec_changed)
        form.addRow("Default vectorizer:", self.vec_combo)

        self.embed_model_edit = QtWidgets.QLineEdit(
            str(self._settings.get("embedding_model", "sentence-transformers/all-MiniLM-L6-v2"))
        )
        self.embed_model_edit.setToolTip(
            "HuggingFace model id used when the vectorizer is set to Embeddings."
        )
        self.embed_model_edit.editingFinished.connect(self._on_embed_model_changed)
        form.addRow("Embedding model:", self.embed_model_edit)

        # Categorization defaults
        self.min_cluster_size_spin = QtWidgets.QSpinBox()
        self.min_cluster_size_spin.setRange(2, 200)
        self.min_cluster_size_spin.setValue(int(self._settings.get("last_min_cluster_size", 5) or 5))
        self.min_cluster_size_spin.setToolTip(
            "Default min_cluster_size for HDBSCAN when Run Categorization opens."
        )
        self.min_cluster_size_spin.valueChanged.connect(self._on_min_cluster_size_changed)
        form.addRow("Default min sub-cluster size:", self.min_cluster_size_spin)

        self.confidence_spin = QtWidgets.QDoubleSpinBox()
        self.confidence_spin.setRange(0.0, 1.0)
        self.confidence_spin.setSingleStep(0.05)
        self.confidence_spin.setDecimals(2)
        self.confidence_spin.setValue(float(self._settings.get("taxonomy_confidence_threshold", 0.45) or 0.45))
        self.confidence_spin.setToolTip(
            "Cosine-similarity cutoff for apply_taxonomy. Rows below this fall "
            "into Non-Repetitive when re-applying a saved taxonomy."
        )
        self.confidence_spin.valueChanged.connect(self._on_confidence_changed)
        form.addRow("Taxonomy confidence threshold:", self.confidence_spin)

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

    def _on_default_vec_changed(self, _idx: int) -> None:
        kind = self.vec_combo.currentData() or "tfidf"
        self._settings["last_vectorizer_kind"] = str(kind)
        app_settings.save(self._settings)
        self.settings_changed.emit({"last_vectorizer_kind": str(kind)})

    def _on_embed_model_changed(self) -> None:
        model = self.embed_model_edit.text().strip() or "sentence-transformers/all-MiniLM-L6-v2"
        self._settings["embedding_model"] = model
        app_settings.save(self._settings)
        self.settings_changed.emit({"embedding_model": model})

    def _on_min_cluster_size_changed(self, value: int) -> None:
        self._settings["last_min_cluster_size"] = int(value)
        app_settings.save(self._settings)
        self.settings_changed.emit({"last_min_cluster_size": int(value)})

    def _on_confidence_changed(self, value: float) -> None:
        self._settings["taxonomy_confidence_threshold"] = float(value)
        app_settings.save(self._settings)
        self.settings_changed.emit({"taxonomy_confidence_threshold": float(value)})

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
