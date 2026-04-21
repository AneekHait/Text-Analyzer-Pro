import os
import sys
import threading
import webbrowser
from datetime import datetime

import joblib
import numpy as np
from PIL.ImageQt import ImageQt
from PySide6 import QtCore, QtGui, QtWidgets

from cluster_tool import (
    TextCleaningConfig,
    assign_cluster_names,
    cluster_texts,
    coerce_text_column,
    get_file_extension,
    get_sheet_names,
    get_top_keywords_per_cluster,
    load_table,
    prepare_text_cleaning,
    save_results,
    vectorize_texts,
    visualize_embeddings,
)
from wordcloud_tool import (
    WordCloudConfig,
    delete_preset,
    deserialize_preset_config,
    export_term_stats,
    get_builtin_shape_names,
    get_color_modes,
    get_default_visual_config,
    get_effective_stopwords,
    get_font_choices,
    get_mask_modes,
    get_palette_names,
    get_sort_modes,
    get_template_config,
    get_template_names,
    load_presets,
    prepare_wordcloud_data,
    render_wordcloud,
    save_preset,
)


def show_error(parent, title, message):
    QtWidgets.QMessageBox.critical(parent, title, str(message))


def show_warning(parent, title, message):
    QtWidgets.QMessageBox.warning(parent, title, str(message))


def show_info(parent, title, message):
    QtWidgets.QMessageBox.information(parent, title, str(message))


class CollapsibleSection(QtWidgets.QWidget):
    def __init__(self, title, subtitle="", expanded=True, parent=None):
        super().__init__(parent)
        self.toggle_button = QtWidgets.QToolButton(text=title)
        self.toggle_button.setCheckable(True)
        self.toggle_button.setChecked(expanded)
        self.toggle_button.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        self.toggle_button.setArrowType(QtCore.Qt.DownArrow if expanded else QtCore.Qt.RightArrow)
        self.subtitle_label = QtWidgets.QLabel(subtitle)
        self.subtitle_label.setWordWrap(True)
        self.content = QtWidgets.QWidget()
        self.content.setVisible(expanded)

        header_layout = QtWidgets.QVBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.addWidget(self.toggle_button)
        if subtitle:
            self.subtitle_label.setStyleSheet("color: #666666;")
            header_layout.addWidget(self.subtitle_label)

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addLayout(header_layout)
        outer.addWidget(self.content)

        self.toggle_button.toggled.connect(self._toggle)

    def _toggle(self, checked):
        self.toggle_button.setArrowType(QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow)
        self.content.setVisible(checked)


class AboutDialog(QtWidgets.QDialog):
    def __init__(self, owner_name, owner_contact, owner_website, owner_bmc, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Text Analyzer Pro")
        self.resize(920, 720)

        text = f"""
Text Analyzer Pro - v1.5

Owner: {owner_name}
LinkedIn: {owner_contact}
Website: {owner_website}
Support: {owner_bmc}

License: MIT
Privacy: All processing happens locally on your machine.

This app supports clustering and wordcloud analysis for Excel, CSV, and JSON files.
"""

        layout = QtWidgets.QVBoxLayout(self)
        header = QtWidgets.QLabel("Text Analyzer Pro — v1.5")
        header.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(header)

        browser = QtWidgets.QPlainTextEdit()
        browser.setReadOnly(True)
        browser.setPlainText(text)
        layout.addWidget(browser, 1)

        button_row = QtWidgets.QHBoxLayout()
        for label, url in (
            ("LinkedIn", owner_contact),
            ("Website", owner_website),
            ("Buy Me a Coffee", owner_bmc),
        ):
            if url:
                button = QtWidgets.QPushButton(label)
                button.clicked.connect(lambda _checked=False, target=url: webbrowser.open(target))
                button_row.addWidget(button)
        button_row.addStretch(1)
        close_button = QtWidgets.QPushButton("Close")
        close_button.clicked.connect(self.accept)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)


class ClusterGUI(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.app_title = "Text Analyzer Pro - v1.5"
        self.owner_name = "Aneek Hait"
        self.owner_contact = "https://www.linkedin.com/in/aneekhait/"
        self.owner_website = "https://aneekhait.github.io"
        self.owner_bmc = "https://buymeacoffee.com/aneekh"

        self.df = None
        self.labels = None
        self.cluster_names = {}
        self.top_keywords = {}
        self.X = None
        self.model = None
        self.vectorizer = None
        self.wordcloud_builder = None
        self.current_file_path = None
        self.latest_cleaning_result = None
        self.cleaning_recipes = {}
        self.name_entries = {}

        self.setWindowTitle(self.app_title)
        self.resize(1180, 860)
        self.setMinimumSize(980, 720)
        self._build_ui()
        self.refresh_cleaning_preview()

    def _build_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)

        menubar = self.menuBar()
        help_menu = menubar.addMenu("Help")
        about_action = QtGui.QAction("About", self)
        about_action.triggered.connect(self.show_imprint)
        help_menu.addAction(about_action)

        self.main_tabs = QtWidgets.QTabWidget()
        root.addWidget(self.main_tabs, 1)

        self.setup_tab = QtWidgets.QWidget()
        self.cleaning_tab = QtWidgets.QWidget()
        self.results_tab = QtWidgets.QWidget()
        self.main_tabs.addTab(self.setup_tab, "Setup")
        self.main_tabs.addTab(self.cleaning_tab, "Cleaning")
        self.main_tabs.addTab(self.results_tab, "Results")

        self._build_setup_tab()
        self._build_cleaning_tab()
        self._build_results_tab()

        action_row = QtWidgets.QHBoxLayout()
        self.run_btn = QtWidgets.QPushButton("Run Clustering")
        self.run_btn.clicked.connect(self.run_clustering_thread)
        self.save_btn = QtWidgets.QPushButton("Save Results")
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self.save_with_names)
        self.vis_btn = QtWidgets.QPushButton("Visualize")
        self.vis_btn.setEnabled(False)
        self.vis_btn.clicked.connect(self.visualize_clusters)
        self.wordcloud_btn = QtWidgets.QPushButton("Generate Wordcloud")
        self.wordcloud_btn.setEnabled(False)
        self.wordcloud_btn.clicked.connect(self.open_wordcloud_builder)
        self.save_model_btn = QtWidgets.QPushButton("Save Model")
        self.save_model_btn.setEnabled(False)
        self.save_model_btn.clicked.connect(self.save_model)
        clear_log_btn = QtWidgets.QPushButton("Clear Log")
        clear_log_btn.clicked.connect(self.clear_log)
        for button in (self.run_btn, self.save_btn, self.vis_btn, self.wordcloud_btn, self.save_model_btn, clear_log_btn):
            action_row.addWidget(button)
        action_row.addStretch(1)
        root.addLayout(action_row)

        self.progress = QtWidgets.QProgressBar()
        self.progress.hide()
        root.addWidget(self.progress)

        footer = QtWidgets.QLabel(f"© 2026 {self.owner_name} • {self.owner_website}")
        footer.setStyleSheet("color: #888888;")
        root.addWidget(footer)

    def _build_setup_tab(self):
        layout = QtWidgets.QVBoxLayout(self.setup_tab)

        file_group = QtWidgets.QGroupBox("File Selection")
        file_layout = QtWidgets.QGridLayout(file_group)
        self.file_label = QtWidgets.QLabel("No file selected")
        self.file_btn = QtWidgets.QPushButton("Select File...")
        self.file_btn.clicked.connect(self.select_file)
        self.sheet_combo = QtWidgets.QComboBox()
        self.sheet_combo.currentTextChanged.connect(self._load_sheet)
        file_layout.addWidget(QtWidgets.QLabel("File:"), 0, 0)
        file_layout.addWidget(self.file_label, 0, 1)
        file_layout.addWidget(self.file_btn, 0, 2)
        file_layout.addWidget(QtWidgets.QLabel("Sheet/Table:"), 1, 0)
        file_layout.addWidget(self.sheet_combo, 1, 1, 1, 2)
        layout.addWidget(file_group)

        params_group = QtWidgets.QGroupBox("Clustering Parameters")
        params_layout = QtWidgets.QGridLayout(params_group)
        self.col_combo = QtWidgets.QComboBox()
        self.col_combo.currentTextChanged.connect(self.refresh_cleaning_preview)
        self.alg_combo = QtWidgets.QComboBox()
        self.alg_combo.addItems(["kmeans", "dbscan", "agglomerative"])
        self.alg_combo.currentTextChanged.connect(self._on_alg_change)
        self.k_spin = QtWidgets.QSpinBox()
        self.k_spin.setRange(1, 999)
        self.k_spin.setValue(5)
        self.name_top_spin = QtWidgets.QSpinBox()
        self.name_top_spin.setRange(1, 50)
        self.name_top_spin.setValue(3)
        self.joiner_edit = QtWidgets.QLineEdit("_")
        self.vis_combo = QtWidgets.QComboBox()
        self.vis_combo.addItems(["pca", "tsne"])
        self.out_edit = QtWidgets.QLineEdit()
        widgets = [
            ("Text column:", self.col_combo),
            ("Algorithm:", self.alg_combo),
            ("n_clusters:", self.k_spin),
            ("name top N:", self.name_top_spin),
            ("joiner:", self.joiner_edit),
            ("Visualization:", self.vis_combo),
        ]
        positions = [(0, 0), (0, 2), (1, 0), (1, 2), (2, 0), (2, 2)]
        for (label, widget), (row, col) in zip(widgets, positions):
            params_layout.addWidget(QtWidgets.QLabel(label), row, col)
            params_layout.addWidget(widget, row, col + 1)
        params_layout.addWidget(QtWidgets.QLabel("Output file:"), 3, 0)
        params_layout.addWidget(self.out_edit, 3, 1, 1, 3)
        layout.addWidget(params_group)
        layout.addStretch(1)

    def _build_cleaning_tab(self):
        layout = QtWidgets.QVBoxLayout(self.cleaning_tab)

        controls_group = QtWidgets.QGroupBox("Text Cleaning")
        controls_layout = QtWidgets.QGridLayout(controls_group)
        self.cleaning_recipe_combo = QtWidgets.QComboBox()
        load_recipe_btn = QtWidgets.QPushButton("Load Recipe")
        load_recipe_btn.clicked.connect(self.load_selected_cleaning_recipe)
        save_recipe_btn = QtWidgets.QPushButton("Save Recipe")
        save_recipe_btn.clicked.connect(self.save_current_cleaning_recipe)
        delete_recipe_btn = QtWidgets.QPushButton("Delete Recipe")
        delete_recipe_btn.clicked.connect(self.delete_selected_cleaning_recipe)
        controls_layout.addWidget(QtWidgets.QLabel("Recipe:"), 0, 0)
        controls_layout.addWidget(self.cleaning_recipe_combo, 0, 1)
        controls_layout.addWidget(load_recipe_btn, 0, 2)
        controls_layout.addWidget(save_recipe_btn, 0, 3)
        controls_layout.addWidget(delete_recipe_btn, 0, 4)

        self.replace_missing_check = QtWidgets.QCheckBox("Replace missing values")
        self.replace_missing_check.setChecked(True)
        self.missing_value_text_edit = QtWidgets.QLineEdit()
        self.trim_whitespace_check = QtWidgets.QCheckBox("Trim whitespace")
        self.trim_whitespace_check.setChecked(True)
        self.lowercase_cleaning_check = QtWidgets.QCheckBox("Lowercase")
        self.lowercase_cleaning_check.setChecked(True)
        self.collapse_whitespace_check = QtWidgets.QCheckBox("Collapse whitespace")
        self.collapse_whitespace_check.setChecked(True)
        self.remove_punctuation_check = QtWidgets.QCheckBox("Remove punctuation")
        self.remove_numbers_check = QtWidgets.QCheckBox("Remove numbers")
        self.remove_urls_check = QtWidgets.QCheckBox("Remove URLs")
        self.remove_emails_check = QtWidgets.QCheckBox("Remove email addresses")
        self.dedupe_cleaned_rows_check = QtWidgets.QCheckBox("Remove duplicate cleaned rows")
        self.regex_pattern_edit = QtWidgets.QLineEdit()
        self.regex_replacement_edit = QtWidgets.QLineEdit()
        refresh_preview_btn = QtWidgets.QPushButton("Refresh Preview")
        refresh_preview_btn.clicked.connect(self.refresh_cleaning_preview)

        toggles = [
            self.replace_missing_check,
            self.trim_whitespace_check,
            self.lowercase_cleaning_check,
            self.collapse_whitespace_check,
            self.remove_punctuation_check,
            self.remove_numbers_check,
            self.remove_urls_check,
            self.remove_emails_check,
            self.dedupe_cleaned_rows_check,
        ]
        for toggle in toggles:
            toggle.toggled.connect(self.refresh_cleaning_preview)
        self.regex_pattern_edit.editingFinished.connect(self.refresh_cleaning_preview)
        self.regex_replacement_edit.editingFinished.connect(self.refresh_cleaning_preview)

        controls_layout.addWidget(self.replace_missing_check, 1, 0, 1, 2)
        controls_layout.addWidget(self.missing_value_text_edit, 1, 2)
        controls_layout.addWidget(self.trim_whitespace_check, 1, 3)
        controls_layout.addWidget(self.lowercase_cleaning_check, 1, 4)
        controls_layout.addWidget(self.collapse_whitespace_check, 2, 0, 1, 2)
        controls_layout.addWidget(self.remove_punctuation_check, 2, 2)
        controls_layout.addWidget(self.remove_numbers_check, 2, 3)
        controls_layout.addWidget(self.remove_urls_check, 2, 4)
        controls_layout.addWidget(self.remove_emails_check, 3, 0, 1, 2)
        controls_layout.addWidget(self.dedupe_cleaned_rows_check, 3, 2, 1, 2)
        controls_layout.addWidget(refresh_preview_btn, 3, 4)
        controls_layout.addWidget(QtWidgets.QLabel("Regex find:"), 4, 0)
        controls_layout.addWidget(self.regex_pattern_edit, 4, 1, 1, 2)
        controls_layout.addWidget(QtWidgets.QLabel("Replace with:"), 4, 3)
        controls_layout.addWidget(self.regex_replacement_edit, 4, 4)
        layout.addWidget(controls_group)

        preview_group = QtWidgets.QGroupBox("Cleaning Preview")
        preview_layout = QtWidgets.QVBoxLayout(preview_group)
        metrics_layout = QtWidgets.QGridLayout()
        self.cleaning_source_label = QtWidgets.QLabel("No source column selected")
        self.cleaning_output_label = QtWidgets.QLabel("No cleaned column yet")
        self.cleaning_rows_label = QtWidgets.QLabel("0")
        self.cleaning_cluster_rows_label = QtWidgets.QLabel("0")
        self.cleaning_empty_rows_label = QtWidgets.QLabel("0")
        self.cleaning_deduped_rows_label = QtWidgets.QLabel("0")
        metrics = [
            ("Source column:", self.cleaning_source_label),
            ("Cleaned column:", self.cleaning_output_label),
            ("Rows before:", self.cleaning_rows_label),
            ("Rows for clustering:", self.cleaning_cluster_rows_label),
            ("Empty after cleaning:", self.cleaning_empty_rows_label),
            ("Deduped rows:", self.cleaning_deduped_rows_label),
        ]
        for index, (label, widget) in enumerate(metrics):
            metrics_layout.addWidget(QtWidgets.QLabel(label), index // 2, (index % 2) * 2)
            metrics_layout.addWidget(widget, index // 2, (index % 2) * 2 + 1)
        preview_layout.addLayout(metrics_layout)

        self.cleaning_preview_table = QtWidgets.QTableWidget(0, 2)
        self.cleaning_preview_table.setHorizontalHeaderLabels(["Original", "Cleaned"])
        self.cleaning_preview_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        preview_layout.addWidget(self.cleaning_preview_table, 1)
        layout.addWidget(preview_group, 1)

    def _build_results_tab(self):
        layout = QtWidgets.QVBoxLayout(self.results_tab)

        log_group = QtWidgets.QGroupBox("Status Log")
        log_layout = QtWidgets.QVBoxLayout(log_group)
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        log_layout.addWidget(self.log)
        layout.addWidget(log_group, 1)

        names_group = QtWidgets.QGroupBox("Edit Cluster Names")
        names_layout = QtWidgets.QVBoxLayout(names_group)
        self.names_scroll = QtWidgets.QScrollArea()
        self.names_scroll.setWidgetResizable(True)
        self.names_container = QtWidgets.QWidget()
        self.names_form = QtWidgets.QFormLayout(self.names_container)
        self.names_scroll.setWidget(self.names_container)
        names_layout.addWidget(self.names_scroll)
        layout.addWidget(names_group)

    def _on_alg_change(self, *_args):
        self.k_spin.setEnabled(self.alg_combo.currentText() != "dbscan")

    def _build_cleaning_config(self):
        return TextCleaningConfig(
            replace_missing=self.replace_missing_check.isChecked(),
            missing_value_text=self.missing_value_text_edit.text(),
            trim_whitespace=self.trim_whitespace_check.isChecked(),
            lowercase=self.lowercase_cleaning_check.isChecked(),
            collapse_whitespace=self.collapse_whitespace_check.isChecked(),
            remove_punctuation=self.remove_punctuation_check.isChecked(),
            remove_numbers=self.remove_numbers_check.isChecked(),
            remove_urls=self.remove_urls_check.isChecked(),
            remove_emails=self.remove_emails_check.isChecked(),
            regex_pattern=self.regex_pattern_edit.text(),
            regex_replacement=self.regex_replacement_edit.text(),
            dedupe_cleaned_rows=self.dedupe_cleaned_rows_check.isChecked(),
        )

    def _cleaning_config_to_dict(self):
        config = self._build_cleaning_config()
        return {
            "replace_missing": config.replace_missing,
            "missing_value_text": config.missing_value_text,
            "trim_whitespace": config.trim_whitespace,
            "lowercase": config.lowercase,
            "collapse_whitespace": config.collapse_whitespace,
            "remove_punctuation": config.remove_punctuation,
            "remove_numbers": config.remove_numbers,
            "remove_urls": config.remove_urls,
            "remove_emails": config.remove_emails,
            "regex_pattern": config.regex_pattern,
            "regex_replacement": config.regex_replacement,
            "dedupe_cleaned_rows": config.dedupe_cleaned_rows,
        }

    def _apply_cleaning_recipe(self, recipe):
        self.replace_missing_check.setChecked(bool(recipe.get("replace_missing", True)))
        self.missing_value_text_edit.setText(recipe.get("missing_value_text", ""))
        self.trim_whitespace_check.setChecked(bool(recipe.get("trim_whitespace", True)))
        self.lowercase_cleaning_check.setChecked(bool(recipe.get("lowercase", True)))
        self.collapse_whitespace_check.setChecked(bool(recipe.get("collapse_whitespace", True)))
        self.remove_punctuation_check.setChecked(bool(recipe.get("remove_punctuation", False)))
        self.remove_numbers_check.setChecked(bool(recipe.get("remove_numbers", False)))
        self.remove_urls_check.setChecked(bool(recipe.get("remove_urls", False)))
        self.remove_emails_check.setChecked(bool(recipe.get("remove_emails", False)))
        self.regex_pattern_edit.setText(recipe.get("regex_pattern", ""))
        self.regex_replacement_edit.setText(recipe.get("regex_replacement", ""))
        self.dedupe_cleaned_rows_check.setChecked(bool(recipe.get("dedupe_cleaned_rows", False)))

    def _refresh_cleaning_recipe_combo(self):
        current = self.cleaning_recipe_combo.currentText()
        self.cleaning_recipe_combo.blockSignals(True)
        self.cleaning_recipe_combo.clear()
        self.cleaning_recipe_combo.addItems(sorted(self.cleaning_recipes.keys()))
        if current in self.cleaning_recipes:
            self.cleaning_recipe_combo.setCurrentText(current)
        self.cleaning_recipe_combo.blockSignals(False)

    def save_current_cleaning_recipe(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Save Recipe", "Recipe name:")
        if not ok:
            return
        clean_name = name.strip()
        if not clean_name:
            show_warning(self, "Missing Name", "Enter a recipe name before saving.")
            return
        self.cleaning_recipes[clean_name] = self._cleaning_config_to_dict()
        self._refresh_cleaning_recipe_combo()
        self.cleaning_recipe_combo.setCurrentText(clean_name)
        self.log_msg(f"✓ Cleaning recipe saved: {clean_name}")

    def load_selected_cleaning_recipe(self):
        recipe_name = self.cleaning_recipe_combo.currentText().strip()
        if not recipe_name:
            show_warning(self, "No Recipe", "Select a cleaning recipe to load.")
            return
        recipe = self.cleaning_recipes.get(recipe_name)
        if recipe is None:
            show_warning(self, "Missing Recipe", f"Recipe '{recipe_name}' is no longer available.")
            self._refresh_cleaning_recipe_combo()
            return
        self._apply_cleaning_recipe(recipe)
        self.refresh_cleaning_preview()
        self.log_msg(f"✓ Cleaning recipe loaded: {recipe_name}")

    def delete_selected_cleaning_recipe(self):
        recipe_name = self.cleaning_recipe_combo.currentText().strip()
        if not recipe_name:
            show_warning(self, "No Recipe", "Select a cleaning recipe to delete.")
            return
        if QtWidgets.QMessageBox.question(self, "Delete Recipe", f"Delete recipe '{recipe_name}'?") != QtWidgets.QMessageBox.Yes:
            return
        self.cleaning_recipes.pop(recipe_name, None)
        self._refresh_cleaning_recipe_combo()
        self.log_msg(f"✓ Cleaning recipe deleted: {recipe_name}")

    def _get_cleaned_column_name(self, source_column):
        return f"{source_column}_cleaned"

    def refresh_cleaning_preview(self):
        self.cleaning_preview_table.setRowCount(0)
        if self.df is None or not self.col_combo.currentText():
            self.cleaning_source_label.setText("No source column selected")
            self.cleaning_output_label.setText("No cleaned column yet")
            self.cleaning_rows_label.setText("0")
            self.cleaning_cluster_rows_label.setText("0")
            self.cleaning_empty_rows_label.setText("0")
            self.cleaning_deduped_rows_label.setText("0")
            return
        try:
            source_column = self.col_combo.currentText().strip()
            result = prepare_text_cleaning(self.df[source_column].tolist(), self._build_cleaning_config())
            self.cleaning_source_label.setText(source_column)
            self.cleaning_output_label.setText(self._get_cleaned_column_name(source_column))
            self.cleaning_rows_label.setText(str(result.stats["source_row_count"]))
            self.cleaning_cluster_rows_label.setText(str(result.stats["kept_row_count"]))
            self.cleaning_empty_rows_label.setText(str(result.stats["empty_row_count"]))
            self.cleaning_deduped_rows_label.setText(str(result.stats["deduped_row_count"]))
            for row in result.preview_rows:
                row_index = self.cleaning_preview_table.rowCount()
                self.cleaning_preview_table.insertRow(row_index)
                self.cleaning_preview_table.setItem(row_index, 0, QtWidgets.QTableWidgetItem(row["raw"]))
                self.cleaning_preview_table.setItem(row_index, 1, QtWidgets.QTableWidgetItem(row["cleaned"]))
        except Exception as error:
            self.log_msg(f"✗ Cleaning preview failed: {error}")

    def status_message(self, message):
        self.statusBar().showMessage(message)

    def log_msg(self, msg):
        formatted = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        self.log.appendPlainText(formatted)

    def clear_log(self):
        self.log.clear()

    def show_imprint(self):
        dialog = AboutDialog(self.owner_name, self.owner_contact, self.owner_website, self.owner_bmc, self)
        dialog.exec()

    def current_column_name(self):
        return self.col_combo.currentText().strip()

    def current_sheet_name(self):
        return self.sheet_combo.currentText().strip()

    def select_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select Input File",
            "",
            "Supported files (*.xlsx *.xls *.csv *.json);;Excel files (*.xlsx *.xls);;CSV files (*.csv);;JSON files (*.json)",
        )
        if not path:
            return
        self.current_file_path = path
        self.file_label.setText(path)
        self.setWindowTitle(f"{self.app_title} - {os.path.basename(path)}")
        base, ext = os.path.splitext(path)
        output_ext = ext if ext in {".xlsx", ".csv", ".json"} else ".xlsx"
        self.out_edit.setText(base + "_clustered" + output_ext)
        try:
            sheet_names = get_sheet_names(path)
            self.sheet_combo.blockSignals(True)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheet_names)
            self.sheet_combo.blockSignals(False)
            if sheet_names:
                self.sheet_combo.setCurrentText(sheet_names[0])
                self._load_sheet(sheet_names[0])
            ext_label = get_file_extension(path)
            if ext_label in {".csv", ".json"}:
                self.log_msg(f"✓ Loaded {ext_label[1:].upper()} source with 1 table: {sheet_names[0]}")
            else:
                self.log_msg(f"✓ Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
        except Exception as error:
            show_error(self, "Error", f"Failed to read file structure: {error}")
            self.log_msg(f"✗ Error reading file structure: {error}")
            self.wordcloud_btn.setEnabled(False)

    def _load_sheet(self, sheet_name):
        if not self.current_file_path or not sheet_name:
            return
        try:
            ext = get_file_extension(self.current_file_path)
            selected_sheet = sheet_name if ext in {".xlsx", ".xls"} else None
            self.df = load_table(self.current_file_path, sheet_name=selected_sheet)
            cols = list(self.df.columns)
            self.col_combo.blockSignals(True)
            self.col_combo.clear()
            self.col_combo.addItems(cols)
            self.col_combo.blockSignals(False)
            if cols:
                self.col_combo.setCurrentText(cols[0])
                self.wordcloud_btn.setEnabled(True)
            else:
                self.wordcloud_btn.setEnabled(False)
            source_label = "sheet" if ext in {".xlsx", ".xls"} else "table"
            file_size_kb = os.path.getsize(self.current_file_path) / 1024
            self.log_msg(f"✓ Loaded {source_label} '{sheet_name}': {len(self.df)} rows, {len(cols)} columns, {file_size_kb:.1f} KB")
            self.log_msg(f"  Columns: {', '.join(cols)}")
            self.refresh_cleaning_preview()
            if self.wordcloud_builder is not None:
                self.wordcloud_builder.refresh_from_app()
            self.main_tabs.setCurrentWidget(self.cleaning_tab)
        except Exception as error:
            show_error(self, "Error", f"Failed to load data: {error}")
            self.log_msg(f"✗ Error loading data: {error}")

    def run_clustering_thread(self):
        self.run_clustering()

    def run_clustering(self):
        if self.df is None:
            show_warning(self, "No file", "Please select an input file first")
            return
        col = self.current_column_name()
        if not col:
            show_warning(self, "No column", "Please select a text column")
            return
        try:
            n_clusters = int(self.k_spin.value())
            if self.alg_combo.currentText() != "dbscan" and n_clusters < 2:
                show_warning(self, "Invalid parameter", "n_clusters must be at least 2")
                return
            top_n = int(self.name_top_spin.value())
        except ValueError as error:
            show_error(self, "Invalid input", f"Please enter valid numbers for parameters: {error}")
            return

        try:
            self.run_btn.setEnabled(False)
            self.save_btn.setEnabled(False)
            self.vis_btn.setEnabled(False)
            self.save_model_btn.setEnabled(False)
            self.progress.show()
            self.progress.setValue(0)

            self.log_msg("=" * 60)
            self.log_msg(f"Starting clustering (Algorithm: {self.alg_combo.currentText()}, n_clusters: {n_clusters})")
            self.progress.setValue(10)
            QtWidgets.QApplication.processEvents()

            cleaning_result = prepare_text_cleaning(self.df[col].tolist(), self._build_cleaning_config())
            cleaned_column_name = self._get_cleaned_column_name(col)
            self.df[cleaned_column_name] = cleaning_result.cleaned_texts
            self.latest_cleaning_result = cleaning_result
            if cleaned_column_name not in [self.col_combo.itemText(i) for i in range(self.col_combo.count())]:
                self.col_combo.addItem(cleaned_column_name)
            self.refresh_cleaning_preview()

            kept_count = cleaning_result.stats["kept_row_count"]
            if not cleaning_result.cluster_input_texts:
                raise ValueError("Cleaning produced no usable text rows. Adjust the cleaning settings and try again.")
            if self.alg_combo.currentText() != "dbscan" and n_clusters > kept_count:
                raise ValueError(f"n_clusters ({n_clusters}) cannot exceed the cleaned clustering row count ({kept_count}).")

            self.progress.setValue(35)
            vectorizer, X = vectorize_texts(cleaning_result.cluster_input_texts)
            self.X = X
            self.vectorizer = vectorizer
            self.log_msg(f"  Vectorizer created: {X.shape[0]} documents, {X.shape[1]} features")

            self.progress.setValue(70)
            self.model, labels = cluster_texts(X, algorithm=self.alg_combo.currentText(), n_clusters=n_clusters)
            label_by_representative = {
                representative_index: int(label)
                for representative_index, label in zip(cleaning_result.kept_indices, labels)
            }
            full_labels = np.full(len(self.df), -1, dtype=int)
            for row_index, representative_index in enumerate(cleaning_result.representative_index_by_row):
                if representative_index is None:
                    continue
                full_labels[row_index] = label_by_representative[representative_index]
            self.labels = full_labels
            self.df["cluster_label"] = full_labels

            self.progress.setValue(90)
            self.top_keywords = get_top_keywords_per_cluster(vectorizer, X, labels, top_n=10)
            self.cluster_names = assign_cluster_names(self.top_keywords, name_top_n=top_n, joiner=self.joiner_edit.text())
            self.populate_name_entries()
            self.main_tabs.setCurrentWidget(self.results_tab)

            self.log_msg("✓ Clustering finished! Edit names below and click 'Save results'")
            self.progress.setValue(100)
            self.save_btn.setEnabled(True)
            self.vis_btn.setEnabled(True)
            self.save_model_btn.setEnabled(True)
        except Exception as error:
            self.log_msg(f"✗ Clustering error: {error}")
            show_error(self, "Error during clustering", error)
        finally:
            self.run_btn.setEnabled(True)
            self.progress.hide()

    def populate_name_entries(self):
        while self.names_form.rowCount():
            self.names_form.removeRow(0)
        self.name_entries = {}
        for cid in sorted(self.cluster_names.keys()):
            entry = QtWidgets.QLineEdit(self.cluster_names[cid])
            keywords = QtWidgets.QLabel(", ".join([term for term, _score in self.top_keywords.get(cid, [])]))
            keywords.setWordWrap(True)
            row_widget = QtWidgets.QWidget()
            row_layout = QtWidgets.QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.addWidget(entry, 1)
            row_layout.addWidget(keywords, 2)
            self.names_form.addRow(f"{cid}:", row_widget)
            self.name_entries[cid] = entry

    def visualize_clusters(self):
        if self.X is None or self.labels is None:
            show_warning(self, "Nothing to visualize", "Run clustering first")
            return
        method = self.vis_combo.currentText()
        self.log_msg(f"Generating {method.upper()} visualization...")
        try:
            visualize_embeddings(self.X, self.labels, method=method)
            self.log_msg(f"✓ {method.upper()} visualization displayed")
        except Exception as error:
            self.log_msg(f"✗ Visualization failed: {error}")
            show_error(self, "Visualization failed", error)

    def open_wordcloud_builder(self):
        if self.df is None:
            show_warning(self, "No file", "Please select a file and table first")
            return
        if self.wordcloud_builder is not None and self.wordcloud_builder.isVisible():
            self.wordcloud_builder.refresh_from_app()
            self.wordcloud_builder.raise_()
            self.wordcloud_builder.activateWindow()
            return
        self.wordcloud_builder = WordCloudBuilderWindow(self)
        self.wordcloud_builder.show()

    def save_with_names(self):
        if self.df is None or self.labels is None:
            show_warning(self, "Nothing to save", "Run clustering first")
            return
        final_names = {}
        for cid, entry in self.name_entries.items():
            name = entry.text().strip()
            if not name:
                show_warning(self, "Invalid input", f"Cluster name for cluster {cid} cannot be empty")
                return
            final_names[cid] = name
        self.df["cluster_name"] = [final_names.get(int(label), "") for label in self.labels]
        out = self.out_edit.text().strip()
        if not out:
            show_warning(self, "No output", "Provide an output filepath")
            return
        try:
            saved_path = save_results(self.df, out)
            self.log_msg(f"✓ Results saved to {saved_path}")
            show_info(self, "Saved", f"Saved results to {saved_path}")
        except Exception as error:
            self.log_msg(f"✗ Save failed: {error}")
            show_error(self, "Save failed", error)

    def save_model(self):
        if self.model is None or self.vectorizer is None:
            show_warning(self, "Nothing to save", "Run clustering first")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Clustering Model", "", "Joblib files (*.joblib)")
        if not path:
            return
        try:
            joblib.dump(
                {
                    "model": self.model,
                    "vectorizer": self.vectorizer,
                    "cluster_names": self.cluster_names,
                    "top_keywords": self.top_keywords,
                },
                path,
            )
            self.log_msg(f"✓ Model saved to {path}")
            show_info(self, "Model Saved", f"Saved model to {path}")
        except Exception as error:
            self.log_msg(f"✗ Model save failed: {error}")
            show_error(self, "Save Failed", f"Failed to save model: {error}")


class WordCloudBuilderWindow(QtWidgets.QDialog):
    render_ready = QtCore.Signal(int, str, object, object, object)
    render_empty = QtCore.Signal(int, str, object)
    render_error = QtCore.Signal(int, str)

    PHRASE_OPTIONS = ("Unigrams", "Up to Bigrams", "Up to Trigrams")
    BACKGROUND_OPTIONS = ("white", "ivory", "whitesmoke", "mintcream", "black", "midnightblue")
    COLORMAP_OPTIONS = ("viridis", "plasma", "inferno", "magma", "cividis", "Set2", "tab10", "cubehelix")
    COLOR_MODE_OPTIONS = get_color_modes()
    MASK_MODE_OPTIONS = get_mask_modes()
    SHAPE_OPTIONS = get_builtin_shape_names()
    PALETTE_OPTIONS = get_palette_names()
    TEMPLATE_OPTIONS = get_template_names()
    SORT_MODE_OPTIONS = get_sort_modes()
    FONT_OPTIONS = get_font_choices()
    LIVE_PREVIEW_DELAY_MS = 450

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.setWindowTitle(f"{self.app.app_title} - Wordcloud Studio")
        self.resize(1360, 860)
        self.setMinimumSize(1180, 760)

        self.custom_stopwords = set()
        self.include_terms = set()
        self.exclude_terms = set()
        self.current_stats_df = None
        self.current_image = None
        self.preview_pixmap = None
        self.is_rendering = False
        self.presets = {}
        self._latest_request_id = 0
        self._pending_preview = False
        self._form_sync_in_progress = False

        self.live_timer = QtCore.QTimer(self)
        self.live_timer.setSingleShot(True)
        self.live_timer.timeout.connect(self.generate_wordcloud_thread)

        self.render_ready.connect(self._finish_render)
        self.render_empty.connect(self._finish_empty_render)
        self.render_error.connect(self._finish_render_error)

        self._build_ui()
        self._load_presets_from_disk()
        self._refresh_preset_combo()
        self.refresh_from_app(reset_preview=False)
        self._apply_config_to_form(get_default_visual_config(), preserve_column=True)
        self._reset_preview_state("Preview updates automatically as you change settings.", clear_summary=True)
        self._bind_live_updates()
        self.update_stopword_count()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def _build_ui(self):
        root = QtWidgets.QVBoxLayout(self)

        header = QtWidgets.QFrame()
        header_layout = QtWidgets.QHBoxLayout(header)
        header_text = QtWidgets.QVBoxLayout()
        title = QtWidgets.QLabel("Wordcloud Studio")
        title.setStyleSheet("font-size: 18px; font-weight: bold;")
        self.context_label = QtWidgets.QLabel("No active sheet")
        header_text.addWidget(title)
        header_text.addWidget(self.context_label)
        header_layout.addLayout(header_text, 1)
        self.active_template_label = QtWidgets.QLabel("Default")
        self.status_label = QtWidgets.QLabel("Ready")
        reset_btn = QtWidgets.QPushButton("Reset")
        reset_btn.clicked.connect(self.reset_to_default_preset)
        meta = QtWidgets.QFormLayout()
        meta.addRow("Template", self.active_template_label)
        meta.addRow("Status", self.status_label)
        meta.addRow("", reset_btn)
        header_layout.addLayout(meta)
        root.addWidget(header)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        root.addWidget(splitter, 1)

        left = QtWidgets.QScrollArea()
        left.setWidgetResizable(True)
        left_container = QtWidgets.QWidget()
        self.left_layout = QtWidgets.QVBoxLayout(left_container)
        left.setWidget(left_container)
        splitter.addWidget(left)

        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        splitter.addWidget(right)
        splitter.setStretchFactor(1, 1)

        self.content_section = CollapsibleSection("Content", "Choose source text and term filters.")
        self.style_section = CollapsibleSection("Style", "Pick templates, fonts, colors, and shapes.")
        self.presets_section = CollapsibleSection("Presets & Templates", "Reuse built-in looks or saved presets.", expanded=False)
        self.insights_section = CollapsibleSection("Insights", "See active filters and quick summary.")
        for section in (self.content_section, self.style_section, self.presets_section, self.insights_section):
            self.left_layout.addWidget(section)
        self.left_layout.addStretch(1)

        self._build_content_section(self.content_section.content)
        self._build_style_section(self.style_section.content)
        self._build_presets_section(self.presets_section.content)
        self._build_left_insights(self.insights_section.content)

        self.validation_label = QtWidgets.QLabel("Adjust a setting to see the preview update automatically.")
        self.validation_label.setWordWrap(True)
        right_layout.addWidget(self.validation_label)

        self.preview_label = QtWidgets.QLabel("Preview updates automatically as you change settings.")
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_label.setMinimumHeight(380)
        self.preview_label.setStyleSheet("background: white; border: 1px solid #cccccc;")
        right_layout.addWidget(self.preview_label, 1)

        insights_row = QtWidgets.QHBoxLayout()
        summary_group = QtWidgets.QGroupBox("Summary")
        summary_form = QtWidgets.QFormLayout(summary_group)
        self.total_rows_label = QtWidgets.QLabel("0")
        self.usable_rows_label = QtWidgets.QLabel("0")
        self.unique_terms_label = QtWidgets.QLabel("0")
        self.term_occurrences_label = QtWidgets.QLabel("0")
        summary_form.addRow("Rows", self.total_rows_label)
        summary_form.addRow("Usable", self.usable_rows_label)
        summary_form.addRow("Terms", self.unique_terms_label)
        summary_form.addRow("Occurrences", self.term_occurrences_label)
        insights_row.addWidget(summary_group)

        terms_group = QtWidgets.QGroupBox("Top Terms")
        terms_layout = QtWidgets.QVBoxLayout(terms_group)
        self.terms_table = QtWidgets.QTableWidget(0, 3)
        self.terms_table.setHorizontalHeaderLabels(["Term", "Count", "Share"])
        self.terms_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        terms_layout.addWidget(self.terms_table)
        insights_row.addWidget(terms_group, 1)
        right_layout.addLayout(insights_row)

        action_row = QtWidgets.QHBoxLayout()
        self.generate_btn = QtWidgets.QPushButton("Refresh Preview")
        self.generate_btn.clicked.connect(self.generate_wordcloud_thread)
        self.save_png_btn = QtWidgets.QPushButton("Save PNG")
        self.save_png_btn.clicked.connect(self.save_png)
        self.save_png_btn.setEnabled(False)
        self.export_terms_btn = QtWidgets.QPushButton("Export Terms")
        self.export_terms_btn.clicked.connect(self.export_terms)
        self.export_terms_btn.setEnabled(False)
        self.copy_settings_btn = QtWidgets.QPushButton("Copy Settings")
        self.copy_settings_btn.clicked.connect(self.copy_current_settings)
        for button in (self.generate_btn, self.save_png_btn, self.export_terms_btn, self.copy_settings_btn):
            action_row.addWidget(button)
        action_row.addStretch(1)
        right_layout.addLayout(action_row)

    def _build_content_section(self, parent):
        layout = QtWidgets.QVBoxLayout(parent)
        form = QtWidgets.QFormLayout()
        self.column_combo = QtWidgets.QComboBox()
        self.column_combo.currentTextChanged.connect(lambda _value: self.refresh_from_app(reset_preview=True))
        self.phrase_combo = QtWidgets.QComboBox()
        self.phrase_combo.addItems(self.PHRASE_OPTIONS)
        self.max_words_spin = QtWidgets.QSpinBox()
        self.max_words_spin.setRange(1, 5000)
        self.max_words_spin.setValue(200)
        self.min_frequency_spin = QtWidgets.QSpinBox()
        self.min_frequency_spin.setRange(1, 9999)
        self.min_frequency_spin.setValue(1)
        form.addRow("Source column", self.column_combo)
        form.addRow("Phrase mode", self.phrase_combo)
        form.addRow("Max words", self.max_words_spin)
        form.addRow("Min frequency", self.min_frequency_spin)
        layout.addLayout(form)

        self.use_builtin_stopwords_check = QtWidgets.QCheckBox("Use built-in stopwords")
        self.use_builtin_stopwords_check.setChecked(True)
        self.lowercase_check = QtWidgets.QCheckBox("Lowercase text before counting")
        self.lowercase_check.setChecked(True)
        self.exclude_numeric_check = QtWidgets.QCheckBox("Ignore numeric-only tokens")
        self.exclude_numeric_check.setChecked(True)
        layout.addWidget(self.use_builtin_stopwords_check)
        layout.addWidget(self.lowercase_check)
        layout.addWidget(self.exclude_numeric_check)

        self.content_advanced_check = QtWidgets.QCheckBox("Show advanced content controls")
        self.content_advanced_check.toggled.connect(lambda checked: self.content_advanced_widget.setVisible(checked))
        layout.addWidget(self.content_advanced_check)

        self.content_advanced_widget = QtWidgets.QWidget()
        adv_layout = QtWidgets.QVBoxLayout(self.content_advanced_widget)
        adv_form = QtWidgets.QFormLayout()
        self.render_top_n_spin = QtWidgets.QSpinBox()
        self.render_top_n_spin.setRange(0, 5000)
        self.sort_mode_combo = QtWidgets.QComboBox()
        self.sort_mode_combo.addItems(self.SORT_MODE_OPTIONS)
        adv_form.addRow("Top terms cap", self.render_top_n_spin)
        adv_form.addRow("Sort preview table", self.sort_mode_combo)
        adv_layout.addLayout(adv_form)
        self.stopword_count_label = QtWidgets.QLabel("Effective stopwords: 0")
        adv_layout.addWidget(self._build_token_editor("Stopwords", "stopword_entry", "stopwords_listbox", self.add_custom_stopwords, self.remove_selected_stopwords, self.clear_custom_stopwords))
        adv_layout.addWidget(self.stopword_count_label)
        adv_layout.addWidget(self._build_token_editor("Include terms", "include_entry", "include_listbox", self.add_include_terms, self.remove_selected_include_terms, self.clear_include_terms))
        adv_layout.addWidget(self._build_token_editor("Exclude terms", "exclude_entry", "exclude_listbox", self.add_exclude_terms, self.remove_selected_exclude_terms, self.clear_exclude_terms))
        self.content_advanced_widget.setVisible(False)
        layout.addWidget(self.content_advanced_widget)

    def _build_style_section(self, parent):
        layout = QtWidgets.QVBoxLayout(parent)
        form = QtWidgets.QFormLayout()
        self.template_combo = QtWidgets.QComboBox()
        self.template_combo.addItems(self.TEMPLATE_OPTIONS)
        self.background_combo = QtWidgets.QComboBox()
        self.background_combo.addItems(self.BACKGROUND_OPTIONS)
        self.color_mode_combo = QtWidgets.QComboBox()
        self.color_mode_combo.addItems(self.COLOR_MODE_OPTIONS)
        self.palette_combo = QtWidgets.QComboBox()
        self.palette_combo.addItems(self.PALETTE_OPTIONS)
        self.font_choice_combo = QtWidgets.QComboBox()
        self.font_choice_combo.addItems([label for label, _path in self.FONT_OPTIONS] + ["Custom file"])
        self.font_label = QtWidgets.QLabel("Default font")
        self.mask_mode_combo = QtWidgets.QComboBox()
        self.mask_mode_combo.addItems(self.MASK_MODE_OPTIONS)
        self.shape_combo = QtWidgets.QComboBox()
        self.shape_combo.addItems(self.SHAPE_OPTIONS)
        self.custom_colors_edit = QtWidgets.QLineEdit()
        form.addRow("Template look", self.template_combo)
        form.addRow("Background", self.background_combo)
        form.addRow("Color mode", self.color_mode_combo)
        form.addRow("Palette", self.palette_combo)
        form.addRow("Font", self.font_choice_combo)
        form.addRow("Selected font", self.font_label)
        form.addRow("Mask mode", self.mask_mode_combo)
        form.addRow("Shape", self.shape_combo)
        form.addRow("Custom colors", self.custom_colors_edit)
        layout.addLayout(form)

        font_row = QtWidgets.QHBoxLayout()
        browse_font_btn = QtWidgets.QPushButton("Browse Font")
        browse_font_btn.clicked.connect(self.browse_font)
        reset_font_btn = QtWidgets.QPushButton("Reset Font")
        reset_font_btn.clicked.connect(self.reset_font)
        font_row.addWidget(browse_font_btn)
        font_row.addWidget(reset_font_btn)
        layout.addLayout(font_row)

        self.style_advanced_check = QtWidgets.QCheckBox("Show advanced style controls")
        self.style_advanced_check.toggled.connect(lambda checked: self.style_advanced_widget.setVisible(checked))
        layout.addWidget(self.style_advanced_check)

        self.style_advanced_widget = QtWidgets.QWidget()
        adv_form = QtWidgets.QFormLayout(self.style_advanced_widget)
        self.mask_path_edit = QtWidgets.QLineEdit()
        browse_mask_btn = QtWidgets.QPushButton("Browse")
        browse_mask_btn.clicked.connect(self.browse_mask)
        clear_mask_btn = QtWidgets.QPushButton("Clear")
        clear_mask_btn.clicked.connect(self.clear_mask)
        mask_row = QtWidgets.QHBoxLayout()
        mask_row.addWidget(self.mask_path_edit, 1)
        mask_row.addWidget(browse_mask_btn)
        mask_row.addWidget(clear_mask_btn)
        mask_row_widget = QtWidgets.QWidget()
        mask_row_widget.setLayout(mask_row)
        self.contour_color_edit = QtWidgets.QLineEdit("#111111")
        self.contour_width_spin = QtWidgets.QSpinBox()
        self.contour_width_spin.setRange(0, 100)
        self.prefer_horizontal_spin = QtWidgets.QDoubleSpinBox()
        self.prefer_horizontal_spin.setRange(0.0, 1.0)
        self.prefer_horizontal_spin.setSingleStep(0.05)
        self.prefer_horizontal_spin.setValue(0.9)
        self.relative_scaling_spin = QtWidgets.QDoubleSpinBox()
        self.relative_scaling_spin.setRange(0.0, 1.0)
        self.relative_scaling_spin.setSingleStep(0.05)
        self.relative_scaling_spin.setValue(0.5)
        self.scale_spin = QtWidgets.QSpinBox()
        self.scale_spin.setRange(1, 20)
        self.scale_spin.setValue(1)
        self.repeat_check = QtWidgets.QCheckBox("Repeat words to fill space")
        self.width_spin = QtWidgets.QSpinBox()
        self.width_spin.setRange(100, 8000)
        self.width_spin.setValue(1200)
        self.height_spin = QtWidgets.QSpinBox()
        self.height_spin.setRange(100, 8000)
        self.height_spin.setValue(700)
        adv_form.addRow("Custom PNG mask", mask_row_widget)
        adv_form.addRow("Contour color", self.contour_color_edit)
        adv_form.addRow("Contour width", self.contour_width_spin)
        adv_form.addRow("Horizontal bias", self.prefer_horizontal_spin)
        adv_form.addRow("Relative scaling", self.relative_scaling_spin)
        adv_form.addRow("Render scale", self.scale_spin)
        adv_form.addRow("", self.repeat_check)
        adv_form.addRow("Canvas width", self.width_spin)
        adv_form.addRow("Canvas height", self.height_spin)
        self.style_advanced_widget.setVisible(False)
        layout.addWidget(self.style_advanced_widget)

    def _build_presets_section(self, parent):
        layout = QtWidgets.QFormLayout(parent)
        template_row = QtWidgets.QHBoxLayout()
        self.template_apply_combo = QtWidgets.QComboBox()
        self.template_apply_combo.addItems(self.TEMPLATE_OPTIONS)
        apply_template_btn = QtWidgets.QPushButton("Apply")
        apply_template_btn.clicked.connect(self.apply_selected_template)
        template_row.addWidget(self.template_apply_combo, 1)
        template_row.addWidget(apply_template_btn)
        template_widget = QtWidgets.QWidget()
        template_widget.setLayout(template_row)
        layout.addRow("Built-in template", template_widget)

        preset_row = QtWidgets.QHBoxLayout()
        self.preset_combo = QtWidgets.QComboBox()
        load_btn = QtWidgets.QPushButton("Load")
        load_btn.clicked.connect(self.load_selected_preset)
        save_btn = QtWidgets.QPushButton("Save as New")
        save_btn.clicked.connect(self.save_current_preset)
        duplicate_btn = QtWidgets.QPushButton("Duplicate")
        duplicate_btn.clicked.connect(self.duplicate_current_preset)
        delete_btn = QtWidgets.QPushButton("Delete")
        delete_btn.clicked.connect(self.delete_selected_preset)
        preset_row.addWidget(self.preset_combo, 1)
        for button in (load_btn, save_btn, duplicate_btn, delete_btn):
            preset_row.addWidget(button)
        preset_widget = QtWidgets.QWidget()
        preset_widget.setLayout(preset_row)
        layout.addRow("Saved preset", preset_widget)

    def _build_left_insights(self, parent):
        layout = QtWidgets.QVBoxLayout(parent)
        self.active_filters_label = QtWidgets.QLabel("No active filters")
        self.active_filters_label.setWordWrap(True)
        layout.addWidget(QtWidgets.QLabel("Active filters"))
        layout.addWidget(self.active_filters_label)
        layout.addWidget(QtWidgets.QLabel("Live preview updates automatically after a brief pause."))

    def _build_token_editor(self, title, entry_attr, list_attr, add_fn, remove_fn, clear_fn):
        group = QtWidgets.QGroupBox(title)
        layout = QtWidgets.QVBoxLayout(group)
        row = QtWidgets.QHBoxLayout()
        entry = QtWidgets.QLineEdit()
        button = QtWidgets.QPushButton("Add")
        button.clicked.connect(add_fn)
        row.addWidget(entry, 1)
        row.addWidget(button)
        layout.addLayout(row)
        listbox = QtWidgets.QListWidget()
        layout.addWidget(listbox)
        actions = QtWidgets.QHBoxLayout()
        remove_btn = QtWidgets.QPushButton("Remove Selected")
        remove_btn.clicked.connect(remove_fn)
        clear_btn = QtWidgets.QPushButton("Clear All")
        clear_btn.clicked.connect(clear_fn)
        actions.addWidget(remove_btn)
        actions.addWidget(clear_btn)
        actions.addStretch(1)
        layout.addLayout(actions)
        setattr(self, entry_attr, entry)
        setattr(self, list_attr, listbox)
        return group

    def _bind_live_updates(self):
        widgets = [
            self.template_combo,
            self.background_combo,
            self.color_mode_combo,
            self.palette_combo,
            self.font_choice_combo,
            self.mask_mode_combo,
            self.shape_combo,
            self.column_combo,
            self.phrase_combo,
            self.sort_mode_combo,
        ]
        for widget in widgets:
            widget.currentTextChanged.connect(self._on_live_setting_changed)
        inputs = [
            self.custom_colors_edit,
            self.mask_path_edit,
            self.contour_color_edit,
        ]
        for widget in inputs:
            widget.textChanged.connect(self._on_live_setting_changed)
        numeric = [
            self.max_words_spin,
            self.min_frequency_spin,
            self.render_top_n_spin,
            self.contour_width_spin,
            self.prefer_horizontal_spin,
            self.relative_scaling_spin,
            self.scale_spin,
            self.width_spin,
            self.height_spin,
        ]
        for widget in numeric:
            widget.valueChanged.connect(self._on_live_setting_changed)
        checks = [
            self.use_builtin_stopwords_check,
            self.lowercase_check,
            self.exclude_numeric_check,
            self.repeat_check,
        ]
        for widget in checks:
            widget.toggled.connect(self._on_live_setting_changed)
        self.font_choice_combo.currentTextChanged.connect(lambda _value: self.apply_selected_font())

    def refresh_from_app(self, reset_preview=True):
        columns = list(self.app.df.columns) if self.app.df is not None else []
        current = self.column_combo.currentText()
        preferred = self.app.current_column_name() if hasattr(self.app, "current_column_name") else ""
        self.column_combo.blockSignals(True)
        self.column_combo.clear()
        self.column_combo.addItems(columns)
        if preferred in columns:
            self.column_combo.setCurrentText(preferred)
        elif current in columns:
            self.column_combo.setCurrentText(current)
        elif columns:
            self.column_combo.setCurrentText(columns[0])
        self.column_combo.blockSignals(False)
        file_name = os.path.basename(self.app.current_file_path) if self.app.current_file_path else "No file"
        sheet_name = self.app.current_sheet_name() if hasattr(self.app, "current_sheet_name") else "No sheet"
        self.context_label.setText(f"File: {file_name}    Sheet: {sheet_name}    Column: {self.column_combo.currentText() or 'No column'}")
        self.generate_btn.setEnabled(bool(columns) and not self.is_rendering)
        if reset_preview:
            self._reset_preview_state("Source changed. Preview will refresh automatically.", clear_summary=True)
            self.schedule_live_preview()

    def schedule_live_preview(self):
        self.status_label.setText("Updating…")
        self.validation_label.setText("Waiting for changes to settle before rerendering.")
        self.live_timer.start(self.LIVE_PREVIEW_DELAY_MS)

    def _on_live_setting_changed(self, *_args):
        if self._form_sync_in_progress:
            return
        self._update_visual_control_states()
        self.update_stopword_count()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def _load_presets_from_disk(self):
        try:
            self.presets = load_presets()
        except Exception as error:
            self.presets = {}
            self.app.log_msg(f"✗ Failed to load wordcloud presets: {error}")

    def _refresh_preset_combo(self):
        current = self.preset_combo.currentText()
        self.preset_combo.blockSignals(True)
        self.preset_combo.clear()
        self.preset_combo.addItems(sorted(self.presets.keys()))
        if current in self.presets:
            self.preset_combo.setCurrentText(current)
        self.preset_combo.blockSignals(False)

    def browse_font(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select Font File", "", "Font files (*.ttf *.otf);;All files (*)")
        if not path:
            return
        self.font_choice_combo.setCurrentText("Custom file")
        self.font_label.setText(os.path.basename(path))
        self._custom_font_path = path
        self.schedule_live_preview()

    def apply_selected_font(self):
        label = self.font_choice_combo.currentText().strip()
        font_map = dict(self.FONT_OPTIONS)
        self._custom_font_path = font_map.get(label, getattr(self, "_custom_font_path", ""))
        self.font_label.setText(label or "Default font")

    def reset_font(self):
        self._custom_font_path = ""
        self.font_choice_combo.setCurrentText("Default font")
        self.font_label.setText("Default font")

    def browse_mask(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select PNG Mask", "", "PNG files (*.png)")
        if not path:
            return
        self.mask_path_edit.setText(path)
        self.mask_mode_combo.setCurrentText("Custom PNG")

    def clear_mask(self):
        self.mask_path_edit.clear()
        if self.mask_mode_combo.currentText() == "Custom PNG":
            self.mask_mode_combo.setCurrentText("None")

    def apply_selected_template(self):
        self._apply_template_to_form(self.template_apply_combo.currentText())

    def _apply_template_to_form(self, template_name):
        try:
            config = get_template_config(template_name)
        except Exception as error:
            self.validation_label.setText(str(error))
            self.status_label.setText("Invalid setting")
            return
        self._apply_config_to_form(config, preserve_column=True)
        self.schedule_live_preview()

    def save_current_preset(self):
        try:
            config = self._build_config(validate_only=True)
        except Exception as error:
            show_error(self, "Invalid Preset", error)
            return
        name, ok = QtWidgets.QInputDialog.getText(self, "Save Preset", "Preset name:")
        if not ok:
            return
        clean_name = name.strip()
        if not clean_name:
            show_warning(self, "Missing Name", "Enter a preset name before saving.")
            return
        try:
            save_preset(clean_name, config)
            self.presets = load_presets()
            self._refresh_preset_combo()
            self.preset_combo.setCurrentText(clean_name)
            self.status_label.setText(f"Preset '{clean_name}' saved.")
            self.validation_label.setText("Preset saved and ready to reuse.")
            self.app.log_msg(f"✓ Wordcloud preset saved: {clean_name}")
        except Exception as error:
            show_error(self, "Preset Save Failed", error)

    def duplicate_current_preset(self):
        base = self.preset_combo.currentText().strip() or self.template_combo.currentText().strip() or "Preset"
        name, ok = QtWidgets.QInputDialog.getText(self, "Duplicate Preset", "New preset name:", text=f"{base} Copy")
        if not ok:
            return
        if name.strip():
            self.preset_combo.setCurrentText(name.strip())
            self.save_current_preset()

    def load_selected_preset(self):
        preset_name = self.preset_combo.currentText().strip()
        if not preset_name:
            show_warning(self, "No Preset", "Select a preset to load.")
            return
        if preset_name not in self.presets:
            show_warning(self, "Missing Preset", f"Preset '{preset_name}' is no longer available.")
            self._load_presets_from_disk()
            self._refresh_preset_combo()
            return
        try:
            config = deserialize_preset_config(self.presets[preset_name])
            self._apply_config_to_form(config, preserve_column=True)
            self.schedule_live_preview()
        except Exception as error:
            show_error(self, "Preset Load Failed", error)

    def delete_selected_preset(self):
        preset_name = self.preset_combo.currentText().strip()
        if not preset_name:
            show_warning(self, "No Preset", "Select a preset to delete.")
            return
        if QtWidgets.QMessageBox.question(self, "Delete Preset", f"Delete preset '{preset_name}'?") != QtWidgets.QMessageBox.Yes:
            return
        try:
            delete_preset(preset_name)
            self.presets = load_presets()
            self._refresh_preset_combo()
        except Exception as error:
            show_error(self, "Preset Delete Failed", error)

    def reset_to_default_preset(self):
        self._apply_config_to_form(get_default_visual_config(), preserve_column=True)
        self.status_label.setText("Ready")
        self.validation_label.setText("Default styling restored. Preview is updating.")
        self.schedule_live_preview()

    def _apply_config_to_form(self, config, preserve_column=False):
        previous_column = self.column_combo.currentText()
        self._form_sync_in_progress = True
        try:
            self.max_words_spin.setValue(config.max_words)
            self.min_frequency_spin.setValue(config.min_frequency)
            self.render_top_n_spin.setValue(config.render_top_n)
            self.sort_mode_combo.setCurrentText(config.sort_mode)
            self.width_spin.setValue(config.width)
            self.height_spin.setValue(config.height)
            self.phrase_combo.setCurrentText(config.phrase_mode)
            self.use_builtin_stopwords_check.setChecked(config.use_builtin_stopwords)
            self.lowercase_check.setChecked(config.lowercase)
            self.exclude_numeric_check.setChecked(config.exclude_numeric)
            self.background_combo.setCurrentText(config.background_color)
            self.template_combo.setCurrentText(config.template_name)
            self.template_apply_combo.setCurrentText(config.template_name)
            self.active_template_label.setText(config.template_name)
            self.color_mode_combo.setCurrentText(config.color_mode)
            self.palette_combo.setCurrentText(config.palette_name)
            self.custom_colors_edit.setText(", ".join(config.custom_colors))
            self.mask_mode_combo.setCurrentText(config.mask_mode)
            self.shape_combo.setCurrentText(config.shape_name)
            self.contour_color_edit.setText(config.contour_color)
            self.contour_width_spin.setValue(config.contour_width)
            self.prefer_horizontal_spin.setValue(config.prefer_horizontal)
            self.relative_scaling_spin.setValue(config.relative_scaling)
            self.scale_spin.setValue(config.scale)
            self.repeat_check.setChecked(config.repeat)
            self.mask_path_edit.setText(config.mask_path)
            matched_font = next((label for label, path in self.FONT_OPTIONS if path == config.font_path), "Custom file" if config.font_path else "Default font")
            self.font_choice_combo.setCurrentText(matched_font)
            self.font_label.setText(config.font_label)
            self._custom_font_path = config.font_path
            self.custom_stopwords = set(config.custom_stopwords)
            self.include_terms = set(config.include_terms)
            self.exclude_terms = set(config.exclude_terms)
            if preserve_column and previous_column:
                self.column_combo.setCurrentText(previous_column)
        finally:
            self._form_sync_in_progress = False
        self._refresh_filter_lists()
        self.update_stopword_count()
        self._update_visual_control_states()
        self._update_active_filter_summary()

    def _update_visual_control_states(self):
        color_mode = self.color_mode_combo.currentText()
        mask_mode = self.mask_mode_combo.currentText()
        self.palette_combo.setEnabled(color_mode == "Palette")
        self.custom_colors_edit.setEnabled(color_mode == "Custom")
        self.shape_combo.setEnabled(mask_mode == "Builtin Shape")
        self.mask_path_edit.setEnabled(mask_mode == "Custom PNG")

    def _refresh_filter_lists(self):
        self.stopwords_listbox.clear()
        self.stopwords_listbox.addItems(sorted(self.custom_stopwords))
        self.include_listbox.clear()
        self.include_listbox.addItems(sorted(self.include_terms))
        self.exclude_listbox.clear()
        self.exclude_listbox.addItems(sorted(self.exclude_terms))
        self._refresh_preset_combo()

    def _selected_items(self, list_widget):
        return [item.text() for item in list_widget.selectedItems()]

    def _add_terms(self, entry, target_set):
        raw = entry.text().strip()
        if not raw:
            return
        target_set.update(item.strip().lower() for item in raw.replace("\n", ",").split(",") if item.strip())
        entry.clear()

    def add_custom_stopwords(self):
        self._add_terms(self.stopword_entry, self.custom_stopwords)
        self._refresh_filter_lists()
        self.update_stopword_count()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def remove_selected_stopwords(self):
        for item in self._selected_items(self.stopwords_listbox):
            self.custom_stopwords.discard(item)
        self._refresh_filter_lists()
        self.update_stopword_count()
        self.schedule_live_preview()

    def clear_custom_stopwords(self):
        self.custom_stopwords.clear()
        self._refresh_filter_lists()
        self.update_stopword_count()
        self.schedule_live_preview()

    def add_include_terms(self):
        self._add_terms(self.include_entry, self.include_terms)
        self._refresh_filter_lists()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def remove_selected_include_terms(self):
        for item in self._selected_items(self.include_listbox):
            self.include_terms.discard(item)
        self._refresh_filter_lists()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def clear_include_terms(self):
        self.include_terms.clear()
        self._refresh_filter_lists()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def add_exclude_terms(self):
        self._add_terms(self.exclude_entry, self.exclude_terms)
        self._refresh_filter_lists()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def remove_selected_exclude_terms(self):
        for item in self._selected_items(self.exclude_listbox):
            self.exclude_terms.discard(item)
        self._refresh_filter_lists()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def clear_exclude_terms(self):
        self.exclude_terms.clear()
        self._refresh_filter_lists()
        self._update_active_filter_summary()
        self.schedule_live_preview()

    def update_stopword_count(self):
        try:
            config = self._build_config(validate_only=True)
            effective_count = len(get_effective_stopwords(config))
        except Exception:
            effective_count = len(self.custom_stopwords)
        self.stopword_count_label.setText(f"Effective stopwords: {effective_count}")

    def _update_active_filter_summary(self):
        chips = []
        chips.append(self.phrase_combo.currentText().replace("Up to ", ""))
        if self.use_builtin_stopwords_check.isChecked():
            chips.append("Built-in stopwords")
        if self.include_terms:
            chips.append(f"Include {len(self.include_terms)}")
        if self.exclude_terms:
            chips.append(f"Exclude {len(self.exclude_terms)}")
        if self.render_top_n_spin.value():
            chips.append(f"Top {self.render_top_n_spin.value()}")
        if self.mask_mode_combo.currentText() == "Builtin Shape":
            chips.append(self.shape_combo.currentText())
        elif self.mask_mode_combo.currentText() == "Custom PNG":
            chips.append("Custom mask")
        if self.color_mode_combo.currentText() == "Palette":
            chips.append(self.palette_combo.currentText())
        elif self.color_mode_combo.currentText() == "Custom":
            chips.append("Custom colors")
        self.active_filters_label.setText(" • ".join(chips) if chips else "No active filters")

    def _build_config(self, validate_only=False):
        column = self.column_combo.currentText().strip()
        if self.app.df is None:
            raise ValueError("Load a sheet before opening the wordcloud builder.")
        if not column:
            raise ValueError("Select a source column for the wordcloud.")
        if column not in self.app.df.columns:
            raise ValueError(f"Column '{column}' is no longer available in the active sheet.")
        return WordCloudConfig(
            max_words=self.max_words_spin.value(),
            min_frequency=self.min_frequency_spin.value(),
            width=self.width_spin.value(),
            height=self.height_spin.value(),
            phrase_mode=self.phrase_combo.currentText(),
            use_builtin_stopwords=self.use_builtin_stopwords_check.isChecked(),
            lowercase=self.lowercase_check.isChecked(),
            exclude_numeric=self.exclude_numeric_check.isChecked(),
            background_color=self.background_combo.currentText(),
            colormap="viridis",
            font_path=getattr(self, "_custom_font_path", ""),
            font_label=self.font_label.text(),
            color_mode=self.color_mode_combo.currentText(),
            palette_name=self.palette_combo.currentText(),
            custom_colors=self._parse_custom_colors(),
            contour_color=self.contour_color_edit.text().strip() or "#111111",
            contour_width=self.contour_width_spin.value(),
            mask_mode=self.mask_mode_combo.currentText(),
            shape_name=self.shape_combo.currentText(),
            mask_path=self.mask_path_edit.text().strip(),
            prefer_horizontal=self.prefer_horizontal_spin.value(),
            repeat=self.repeat_check.isChecked(),
            relative_scaling=self.relative_scaling_spin.value(),
            scale=self.scale_spin.value(),
            custom_stopwords=set(self.custom_stopwords),
            include_terms=set(self.include_terms),
            exclude_terms=set(self.exclude_terms),
            render_top_n=self.render_top_n_spin.value(),
            sort_mode=self.sort_mode_combo.currentText(),
            template_name=self.template_combo.currentText(),
        )

    def _parse_custom_colors(self):
        raw = self.custom_colors_edit.text().strip()
        return [item.strip() for item in raw.split(",") if item.strip()] if raw else []

    def generate_wordcloud_thread(self):
        if self.live_timer.isActive():
            self.live_timer.stop()
        try:
            config = self._build_config()
            column = self.column_combo.currentText().strip()
            texts = coerce_text_column(self.app.df[column]).tolist()
        except Exception as error:
            self.status_label.setText("Invalid setting")
            self.validation_label.setText(str(error))
            return
        if self.is_rendering:
            self._pending_preview = True
            return
        self.is_rendering = True
        self.generate_btn.setEnabled(False)
        self.save_png_btn.setEnabled(False)
        self.export_terms_btn.setEnabled(False)
        self.status_label.setText("Updating…")
        self.validation_label.setText("Rebuilding the preview with your latest settings.")
        self._latest_request_id += 1
        request_id = self._latest_request_id
        threading.Thread(target=self._render_worker, args=(request_id, column, texts, config), daemon=True).start()

    def _render_worker(self, request_id, column, texts, config):
        try:
            stats_df, summary = prepare_wordcloud_data(texts, config)
            if stats_df.empty:
                self.render_empty.emit(request_id, column, summary)
                return
            image = render_wordcloud(stats_df, config)
            self.render_ready.emit(request_id, column, stats_df, summary, image)
        except Exception as error:
            self.render_error.emit(request_id, str(error))

    @QtCore.Slot(int, str, object, object, object)
    def _finish_render(self, request_id, column, stats_df, summary, image):
        if request_id != self._latest_request_id:
            self._finalize_render_cycle(stale=True)
            return
        self.current_stats_df = stats_df
        self.current_image = image
        self._update_summary(summary)
        self._populate_terms_table(stats_df)
        self._update_preview_image(image)
        self.status_label.setText("Ready")
        self.validation_label.setText(f"Preview ready for column '{column}'.")
        self.save_png_btn.setEnabled(True)
        self.export_terms_btn.setEnabled(True)
        self._finalize_render_cycle()

    @QtCore.Slot(int, str, object)
    def _finish_empty_render(self, request_id, _column, summary):
        if request_id != self._latest_request_id:
            self._finalize_render_cycle(stale=True)
            return
        self._update_summary(summary)
        self._populate_terms_table(None)
        self.status_label.setText("No terms found")
        self.validation_label.setText("No terms remained after applying the current filters.")
        self._finalize_render_cycle()

    @QtCore.Slot(int, str)
    def _finish_render_error(self, request_id, error_message):
        if request_id != self._latest_request_id:
            self._finalize_render_cycle(stale=True)
            return
        self.status_label.setText("Invalid setting")
        self.validation_label.setText(error_message)
        self._finalize_render_cycle()

    def _finalize_render_cycle(self, stale=False):
        self.is_rendering = False
        self.generate_btn.setEnabled(True)
        if stale or self._pending_preview:
            self._pending_preview = False
            self.schedule_live_preview()

    def _update_summary(self, summary):
        self.total_rows_label.setText(str(summary.get("total_rows", 0)))
        self.usable_rows_label.setText(str(summary.get("usable_rows", 0)))
        self.unique_terms_label.setText(str(summary.get("unique_terms", 0)))
        self.term_occurrences_label.setText(str(summary.get("kept_term_occurrences", 0)))

    def _populate_terms_table(self, stats_df):
        self.terms_table.setRowCount(0)
        if stats_df is None or stats_df.empty:
            return
        for _, row in stats_df.head(12).iterrows():
            idx = self.terms_table.rowCount()
            self.terms_table.insertRow(idx)
            self.terms_table.setItem(idx, 0, QtWidgets.QTableWidgetItem(str(row["term"])))
            self.terms_table.setItem(idx, 1, QtWidgets.QTableWidgetItem(str(int(row["count"]))))
            self.terms_table.setItem(idx, 2, QtWidgets.QTableWidgetItem(f"{row['share']:.1%}"))

    def _update_preview_image(self, image):
        self.current_image = image
        qimage = ImageQt(image)
        pixmap = QtGui.QPixmap.fromImage(qimage)
        self.preview_pixmap = pixmap.scaled(900, 620, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.preview_label.setPixmap(self.preview_pixmap)
        self.preview_label.setText("")

    def _reset_preview_state(self, message, clear_summary=False):
        self.current_stats_df = None
        self.current_image = None
        self.preview_pixmap = None
        self.preview_label.setPixmap(QtGui.QPixmap())
        self.preview_label.setText(message)
        self.save_png_btn.setEnabled(False)
        self.export_terms_btn.setEnabled(False)
        self._populate_terms_table(None)
        if clear_summary:
            self._update_summary({})

    def save_png(self):
        if self.current_image is None:
            show_warning(self, "No Preview", "Generate a wordcloud preview first")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Wordcloud PNG", self._default_export_stem() + ".png", "PNG files (*.png)")
        if not path:
            return
        try:
            self.current_image.save(path, format="PNG")
            self.app.log_msg(f"✓ Wordcloud image saved to {path}")
            show_info(self, "Saved", f"Saved wordcloud image to {path}")
        except Exception as error:
            show_error(self, "Save Failed", f"Failed to save wordcloud image: {error}")

    def export_terms(self):
        if self.current_stats_df is None or self.current_stats_df.empty:
            show_warning(self, "No Terms", "Generate a wordcloud preview first")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export Term Statistics", self._default_export_stem() + "_terms.xlsx", "Excel files (*.xlsx)")
        if not path:
            return
        try:
            saved_path = export_term_stats(self.current_stats_df, path)
            self.app.log_msg(f"✓ Wordcloud terms exported to {saved_path}")
            show_info(self, "Exported", f"Exported wordcloud terms to {saved_path}")
        except Exception as error:
            show_error(self, "Export Failed", f"Failed to export term statistics: {error}")

    def copy_current_settings(self):
        try:
            config = self._build_config(validate_only=True)
        except Exception as error:
            show_warning(self, "Invalid Settings", error)
            return
        text = "\n".join(
            [
                f"Template: {config.template_name}",
                f"Column: {self.column_combo.currentText().strip()}",
                f"Phrase mode: {config.phrase_mode}",
                f"Max words: {config.max_words}",
                f"Min frequency: {config.min_frequency}",
                f"Top cap: {config.render_top_n}",
                f"Color mode: {config.color_mode}",
                f"Mask: {config.mask_mode}",
            ]
        )
        QtWidgets.QApplication.clipboard().setText(text)
        self.validation_label.setText("Current settings copied to the clipboard.")

    def _default_export_stem(self):
        base = os.path.splitext(os.path.basename(self.app.current_file_path or "wordcloud"))[0]
        sheet = self._slugify(self.app.current_sheet_name() if hasattr(self.app, "current_sheet_name") else "sheet")
        column = self._slugify(self.column_combo.currentText() or "column")
        return f"{base}_{sheet}_{column}_wordcloud"

    def _slugify(self, value):
        chars = [char.lower() if char.isalnum() else "_" for char in value]
        slug = "".join(chars).strip("_")
        while "__" in slug:
            slug = slug.replace("__", "_")
        return slug or "item"


def main():
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    app.setApplicationName("Text Analyzer Pro")
    window = ClusterGUI()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
