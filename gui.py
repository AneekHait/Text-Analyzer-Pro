"""Text Analyzer Pro — main window shell.

Reusable helpers, workers and standalone widgets have been split into the
`textanalyzer/` package; this module now hosts only the main `ClusterGUI`
window and the `main()` entry point. A backward-compat re-export of
`WordCloudBuilderWindow` keeps existing test imports working.
"""

import os
import sys
import webbrowser
from datetime import datetime
from pathlib import Path

import numpy as np
from PySide6 import QtCore, QtGui, QtWidgets

import app_settings
from theme import ThemeManager

from cluster_tool import (
    TextCleaningConfig,
    coerce_text_column,
    get_file_extension,
    get_sheet_names,
    load_table,
    prepare_text_cleaning,
    save_results,
    visualize_embeddings,
)

from textanalyzer.controllers.analysis_controller import AnalysisController
from textanalyzer.models.config import CleaningConfigModel
from textanalyzer.services.io import IOService

from textanalyzer.utils.dialogs import show_error, show_info, show_warning
from textanalyzer.utils.icons import (
    ASSETS_DIR,
    _BANNER_CACHE,
    _HAS_QTA,
    _icon,
    _load_app_icon,
    _load_banner_pixmap,
    _spin_icon,
)
from textanalyzer.ui.collapsible import CollapsibleSection
from textanalyzer.ui.data_source_panel import DataSourcePanel
from textanalyzer.ui.embedding_canvas import EmbeddingCanvas, _HAS_MPL
from textanalyzer.ui.sidebar import SidebarButton as _SidebarButton
from textanalyzer.ui.shell.dock_panels import InspectorDock, NavigatorDock
from textanalyzer.ui.shell.workspace_tabs import WorkspaceTabWidget
from textanalyzer.ui.toast import Toast, show_toast
# AboutDialog, DiagnosticsWindow, SettingsWindow, WordCloudDialog are
# lazy-imported inside the methods that open them — keeps their (and their
# transitive) imports off the app launch path. WordCloudDialog in
# particular pulls in the wordcloud package + PIL ImageQt.
from textanalyzer.workers.cluster import ClusterWorker
from textanalyzer.workers.embedding import EmbeddingWorker


def __getattr__(name):
    """Lazy module-level attribute resolution (PEP 562).

    Keeps the legacy ``from gui import WordCloudBuilderWindow`` /
    ``WordCloudDialog`` imports working without forcing
    ``textanalyzer.ui.wordcloud_window`` (and its transitive `wordcloud` +
    `PIL.ImageQt` deps) into the app launch path. The dialog module is
    imported on first attribute access — typically the test harness or an
    external script.
    """
    if name in ("WordCloudDialog", "WordCloudBuilderWindow"):
        from textanalyzer.ui.wordcloud_window import WordCloudDialog as _wcd
        return _wcd
    if name == "AboutDialog":
        from textanalyzer.ui.about import AboutDialog as _ad
        return _ad
    if name == "SettingsWindow":
        from textanalyzer.ui.settings_window import SettingsWindow as _sw
        return _sw
    if name == "DiagnosticsWindow":
        from textanalyzer.ui.diagnostics_window import DiagnosticsWindow as _dw
        return _dw
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")

class ClusterGUI(QtWidgets.QMainWindow):
    PAGE_SETUP = 0
    PAGE_CLEANING = 1
    PAGE_RESULTS = 2

    def __init__(self, theme_manager: "ThemeManager" = None):
        super().__init__()
        self.app_title = "Text Analyzer Pro"
        self.app_version = "v2.0"
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
        self.settings_window: SettingsWindow | None = None
        self.diagnostics_window: DiagnosticsWindow | None = None
        self.current_file_path = None
        self.latest_cleaning_result = None
        self.cleaning_recipes = {}
        self.name_entries = {}

        self._cluster_thread = None
        self._cluster_worker = None
        self._embedding_thread = None

        # Advanced TF-IDF / clustering settings exposed via the Setup tab's
        # "Advanced…" dialog. Sensible defaults match the previous hardcoded
        # behavior so unmodified users see no change.
        self._tfidf_settings = {
            "min_df": 1,
            "max_df": 1.0,
            "ngram_range": (1, 1),
            "use_hashing": False,
            "max_features": 2000,
        }
        self._embedding_worker = None

        self.theme_manager = theme_manager
        self.settings = app_settings.load()
        self.cleaning_recipes = dict(self.settings.get("cleaning_recipes", {}) or {})

        # Controller — owns session model & orchestrates workers/services.
        self.controller = AnalysisController(self.settings, parent=self)
        self._connect_controller_signals()

        self.setWindowTitle(f"{self.app_title} {self.app_version}")
        self.setWindowIcon(_load_app_icon())
        self.resize(1280, 880)
        self.setMinimumSize(1040, 740)

        self._build_ui()
        self._restore_geometry()
        self._refresh_recent_files_menu()
        self._refresh_cleaning_recipe_combo()
        self._select_page(self.PAGE_SETUP)
        self.refresh_cleaning_preview()

    def _connect_controller_signals(self) -> None:
        """Subscribe to AnalysisController signals for view updates."""
        c = self.controller
        c.log.connect(self.log_msg)
        c.cluster_started.connect(lambda: self._set_running_state(True))
        c.cluster_progress.connect(self._on_cluster_progress)
        c.cluster_finished.connect(self._on_ctrl_cluster_finished)
        c.cluster_failed.connect(self._on_ctrl_cluster_failed)
        c.cluster_cancelled.connect(
            lambda: (show_toast(self, "Clustering cancelled", level="warning"),
                     self._set_running_state(False)))
        c.embedding_ready.connect(self._on_ctrl_embedding_ready)
        c.embedding_failed.connect(
            lambda msg: (self.log_msg(f"\u2717 Visualization failed: {msg}"),
                         show_error(self, "Visualization failed", msg)))
        c.results_saved.connect(
            lambda p: show_toast(self, f"Saved to {os.path.basename(p)}", level="success"))
        c.model_saved.connect(
            lambda p: show_toast(self, f"Model saved: {os.path.basename(p)}", level="success"))

    # ------------------------------------------------------------------ #
    # Shell construction                                                 #
    # ------------------------------------------------------------------ #

    def _build_ui(self):
        # Build the menu first so docks can be wired into the View menu.
        self._build_menu()

        # The workspace tab widget is the central document container.
        # Phase 2 wraps the existing single-page UI as the only tab; Phase 3
        # will allow opening multiple analysis sessions in parallel.
        self.workspace_tabs = WorkspaceTabWidget()
        self.setCentralWidget(self.workspace_tabs)

        analysis_tab = QtWidgets.QWidget()
        analysis_tab.setObjectName("AnalysisTab")
        root = QtWidgets.QHBoxLayout(analysis_tab)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self._build_sidebar(root)

        body = QtWidgets.QWidget()
        body_layout = QtWidgets.QVBoxLayout(body)
        body_layout.setContentsMargins(0, 0, 0, 0)
        body_layout.setSpacing(0)
        root.addWidget(body, 1)

        self._build_topbar(body_layout)

        self.stack = QtWidgets.QStackedWidget()
        body_layout.addWidget(self.stack, 1)

        self.setup_page = QtWidgets.QWidget()
        self.cleaning_page = QtWidgets.QWidget()
        self.results_page = QtWidgets.QWidget()
        self.stack.addWidget(self._wrap_page(self.setup_page))
        self.stack.addWidget(self._wrap_page(self.cleaning_page))
        self.stack.addWidget(self._wrap_page(self.results_page))

        self._build_setup_tab()
        self._build_cleaning_tab()
        self._build_results_tab()

        self._build_action_bar(body_layout)
        self._build_status_bar()

        # Add the wrapped legacy UI as the first workspace tab.
        self.workspace_tabs.addTab(analysis_tab, "Analysis")
        # Single tab — hide its close button so it can't be closed away.
        bar = self.workspace_tabs.tabBar()
        for side in (QtWidgets.QTabBar.ButtonPosition.RightSide,
                     QtWidgets.QTabBar.ButtonPosition.LeftSide):
            btn = bar.tabButton(0, side)
            if btn is not None:
                btn.resize(0, 0)
                btn.hide()

        # Dock panels (Navigator left, Inspector right). Hidden by default
        # to preserve the familiar layout; toggleable via the View menu.
        self._build_dock_panels()

    def _wrap_page(self, page_widget: QtWidgets.QWidget) -> QtWidgets.QWidget:
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        container = QtWidgets.QWidget()
        outer = QtWidgets.QVBoxLayout(container)
        outer.setContentsMargins(28, 22, 28, 18)
        outer.setSpacing(14)
        outer.addWidget(page_widget)
        scroll.setWidget(container)
        return scroll

    def _build_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("&File")
        open_action = QtGui.QAction(_icon("fa5s.folder-open"), "&Open File…", self)
        open_action.setShortcut(QtGui.QKeySequence.Open)
        open_action.triggered.connect(self.select_file)
        file_menu.addAction(open_action)

        self.recent_menu = file_menu.addMenu("Open &Recent")
        file_menu.addSeparator()

        save_results_action = QtGui.QAction(_icon("fa5s.save"), "&Save Results", self)
        save_results_action.setShortcut(QtGui.QKeySequence.Save)
        save_results_action.triggered.connect(self.save_with_names)
        file_menu.addAction(save_results_action)

        save_model_action = QtGui.QAction("Save &Model…", self)
        save_model_action.triggered.connect(self.save_model)
        file_menu.addAction(save_model_action)

        load_model_action = QtGui.QAction(_icon("fa5s.upload"), "&Load Model…", self)
        load_model_action.triggered.connect(self.load_model_apply)
        file_menu.addAction(load_model_action)

        file_menu.addSeparator()
        quit_action = QtGui.QAction("E&xit", self)
        quit_action.setShortcut(QtGui.QKeySequence.Quit)
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        run_menu = menubar.addMenu("&Run")
        run_action = QtGui.QAction(_icon("fa5s.play"), "&Run Clustering", self)
        run_action.setShortcut("Ctrl+R")
        run_action.triggered.connect(self.run_clustering_thread)
        run_menu.addAction(run_action)
        viz_action = QtGui.QAction(_icon("fa5s.chart-scatter", color="#3d6dff"), "&Visualize", self)
        viz_action.triggered.connect(self.visualize_clusters)
        run_menu.addAction(viz_action)
        wc_action = QtGui.QAction(_icon("fa5s.cloud"), "Wordcloud &Studio…", self)
        wc_action.triggered.connect(self.open_wordcloud_builder)
        run_menu.addAction(wc_action)

        view_menu = menubar.addMenu("&View")
        self.theme_action = QtGui.QAction("Toggle &Theme", self)
        self.theme_action.setShortcut("Ctrl+T")
        self.theme_action.triggered.connect(self.toggle_theme)
        view_menu.addAction(self.theme_action)

        prefs_action = QtGui.QAction(_icon("fa5s.cog"), "&Preferences…", self)
        prefs_action.setShortcut(QtGui.QKeySequence("Ctrl+,"))
        prefs_action.setMenuRole(QtGui.QAction.MenuRole.PreferencesRole)
        prefs_action.triggered.connect(self.show_settings_window)
        view_menu.addAction(prefs_action)

        diag_action = QtGui.QAction(_icon("fa5s.heartbeat"), "&Diagnostics…", self)
        diag_action.setShortcut(QtGui.QKeySequence("Ctrl+/"))
        diag_action.triggered.connect(self.show_diagnostics_window)
        view_menu.addAction(diag_action)

        view_menu.addSeparator()
        # Dock toggle actions are populated in _build_dock_panels.
        self._view_menu = view_menu

        window_menu = menubar.addMenu("&Window")
        next_tab = QtGui.QAction("&Next Tab", self)
        next_tab.setShortcut(QtGui.QKeySequence("Ctrl+Tab"))
        next_tab.triggered.connect(self._activate_next_tab)
        window_menu.addAction(next_tab)
        prev_tab = QtGui.QAction("&Previous Tab", self)
        prev_tab.setShortcut(QtGui.QKeySequence("Ctrl+Shift+Tab"))
        prev_tab.triggered.connect(self._activate_prev_tab)
        window_menu.addAction(prev_tab)
        close_tab = QtGui.QAction("&Close Tab", self)
        close_tab.setShortcut(QtGui.QKeySequence.StandardKey.Close)
        close_tab.triggered.connect(self._close_active_tab)
        window_menu.addAction(close_tab)

        help_menu = menubar.addMenu("&Help")
        about_action = QtGui.QAction(_icon("fa5s.info-circle"), "&About", self)
        about_action.setShortcut(QtGui.QKeySequence("F1"))
        about_action.setMenuRole(QtGui.QAction.MenuRole.AboutRole)
        about_action.triggered.connect(self.show_imprint)
        help_menu.addAction(about_action)

    def _build_sidebar(self, root_layout: QtWidgets.QHBoxLayout):
        sidebar = QtWidgets.QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(228)
        layout = QtWidgets.QVBoxLayout(sidebar)
        layout.setContentsMargins(14, 18, 14, 14)
        layout.setSpacing(6)

        # Brand row
        banner_pixmap = _load_banner_pixmap(48)
        if not banner_pixmap.isNull():
            banner_label = QtWidgets.QLabel()
            banner_label.setPixmap(banner_pixmap)
            banner_label.setAlignment(QtCore.Qt.AlignCenter)
            layout.addWidget(banner_label)
            version_label = QtWidgets.QLabel(self.app_version)
            version_label.setObjectName("SidebarTagline")
            version_label.setAlignment(QtCore.Qt.AlignCenter)
            layout.addWidget(version_label)
        else:
            brand_row = QtWidgets.QHBoxLayout()
            brand_row.setSpacing(10)
            logo_label = QtWidgets.QLabel()
            app_icon = _load_app_icon()
            if not app_icon.isNull():
                logo_label.setPixmap(app_icon.pixmap(28, 28))
            brand_text = QtWidgets.QVBoxLayout()
            brand_text.setSpacing(0)
            brand_name = QtWidgets.QLabel("Text Analyzer")
            brand_name.setObjectName("SidebarBrand")
            brand_tag = QtWidgets.QLabel("Pro " + self.app_version)
            brand_tag.setObjectName("SidebarTagline")
            brand_text.addWidget(brand_name)
            brand_text.addWidget(brand_tag)
            brand_row.addWidget(logo_label)
            brand_row.addLayout(brand_text, 1)
            layout.addLayout(brand_row)

        divider = QtWidgets.QFrame()
        divider.setObjectName("SidebarDivider")
        layout.addWidget(divider)

        # Nav buttons
        self.nav_buttons = []
        nav_specs = [
            ("Setup", "fa5s.sliders-h", self.PAGE_SETUP),
            ("Cleaning", "fa5s.broom", self.PAGE_CLEANING),
            ("Results", "fa5s.chart-bar", self.PAGE_RESULTS),
        ]
        for label, icon_name, page_index in nav_specs:
            btn = _SidebarButton(label, icon_name)
            btn.clicked.connect(lambda _checked=False, idx=page_index: self._select_page(idx))
            layout.addWidget(btn)
            self.nav_buttons.append(btn)

        layout.addSpacing(8)
        sub_divider = QtWidgets.QFrame()
        sub_divider.setObjectName("SidebarDivider")
        layout.addWidget(sub_divider)

        wc_btn = _SidebarButton("Wordcloud Studio", "fa5s.cloud")
        wc_btn.clicked.connect(self.open_wordcloud_builder)
        layout.addWidget(wc_btn)
        self.sidebar_wordcloud_btn = wc_btn

        viz_btn = _SidebarButton("Visualize", "fa5s.chart-scatter")
        viz_btn.clicked.connect(self.visualize_clusters)
        layout.addWidget(viz_btn)
        self.sidebar_viz_btn = viz_btn

        layout.addStretch(1)

        bottom_divider = QtWidgets.QFrame()
        bottom_divider.setObjectName("SidebarDivider")
        layout.addWidget(bottom_divider)

        theme_btn = _SidebarButton("Toggle Theme", "fa5s.adjust")
        theme_btn.clicked.connect(self.toggle_theme)
        layout.addWidget(theme_btn)

        about_btn = _SidebarButton("About", "fa5s.info-circle")
        about_btn.clicked.connect(self.show_imprint)
        layout.addWidget(about_btn)

        root_layout.addWidget(sidebar)

    def _build_topbar(self, parent_layout: QtWidgets.QVBoxLayout):
        bar = QtWidgets.QFrame()
        bar.setObjectName("TopBar")
        layout = QtWidgets.QHBoxLayout(bar)
        layout.setContentsMargins(28, 14, 28, 14)
        layout.setSpacing(12)

        title_col = QtWidgets.QVBoxLayout()
        title_col.setSpacing(2)
        self.page_title_label = QtWidgets.QLabel("Setup")
        self.page_title_label.setObjectName("PageTitle")
        self.page_subtitle_label = QtWidgets.QLabel("Choose your data and clustering parameters.")
        self.page_subtitle_label.setObjectName("PageSubtitle")
        title_col.addWidget(self.page_title_label)
        title_col.addWidget(self.page_subtitle_label)
        layout.addLayout(title_col, 1)

        self.topbar_file_label = QtWidgets.QLabel("No file loaded")
        self.topbar_file_label.setProperty("role", "muted")
        layout.addWidget(self.topbar_file_label)

        parent_layout.addWidget(bar)

    def _build_action_bar(self, parent_layout: QtWidgets.QVBoxLayout):
        bar = QtWidgets.QFrame()
        bar.setObjectName("ActionBar")
        layout = QtWidgets.QHBoxLayout(bar)
        layout.setContentsMargins(28, 12, 28, 12)
        layout.setSpacing(10)

        self.progress = QtWidgets.QProgressBar()
        self.progress.setMaximumWidth(260)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.hide()
        layout.addWidget(self.progress)

        self.progress_label = QtWidgets.QLabel("")
        self.progress_label.setProperty("role", "muted")
        layout.addWidget(self.progress_label)

        layout.addStretch(1)

        self.cancel_btn = QtWidgets.QPushButton("Cancel")
        self.cancel_btn.setIcon(_icon("fa5s.stop"))
        self.cancel_btn.setProperty("danger", "true")
        self.cancel_btn.clicked.connect(self.cancel_clustering)
        self.cancel_btn.hide()
        layout.addWidget(self.cancel_btn)

        clear_log_btn = QtWidgets.QPushButton("Clear Log")
        clear_log_btn.setIcon(_icon("fa5s.eraser"))
        clear_log_btn.clicked.connect(self.clear_log)
        layout.addWidget(clear_log_btn)

        self.save_model_btn = QtWidgets.QPushButton("Save Model")
        self.save_model_btn.setIcon(_icon("fa5s.cube"))
        self.save_model_btn.setEnabled(False)
        self.save_model_btn.clicked.connect(self.save_model)
        layout.addWidget(self.save_model_btn)

        self.save_btn = QtWidgets.QPushButton("Save Results")
        self.save_btn.setIcon(_icon("fa5s.save"))
        self.save_btn.setEnabled(False)
        self.save_btn.clicked.connect(self.save_with_names)
        layout.addWidget(self.save_btn)

        self.run_btn = QtWidgets.QPushButton("Run Clustering")
        self.run_btn.setIcon(_icon("fa5s.play", color="#ffffff"))
        self.run_btn.setProperty("primary", "true")
        self.run_btn.setShortcut("Ctrl+R")
        self.run_btn.clicked.connect(self.run_clustering_thread)
        layout.addWidget(self.run_btn)

        parent_layout.addWidget(bar)

    def _build_status_bar(self):
        status = self.statusBar()
        self.status_file_label = QtWidgets.QLabel("No file")
        self.status_rows_label = QtWidgets.QLabel("0 rows")
        self.status_col_label = QtWidgets.QLabel("No column")
        self.status_alg_label = QtWidgets.QLabel("kmeans")
        for label in (self.status_file_label, self.status_rows_label, self.status_col_label, self.status_alg_label):
            label.setProperty("role", "muted")
        status.addWidget(self.status_file_label)
        status.addPermanentWidget(self.status_col_label)
        status.addPermanentWidget(self.status_rows_label)
        status.addPermanentWidget(self.status_alg_label)
        self._update_status_bar()

    # ------------------------------------------------------------------ #
    # Shell: dock panels and tab navigation                              #
    # ------------------------------------------------------------------ #

    def _build_dock_panels(self):
        self.navigator_dock = NavigatorDock(self)
        self.inspector_dock = InspectorDock(self)
        self.addDockWidget(QtCore.Qt.DockWidgetArea.LeftDockWidgetArea, self.navigator_dock)
        self.addDockWidget(QtCore.Qt.DockWidgetArea.RightDockWidgetArea, self.inspector_dock)

        # Hidden by default — preserve the legacy layout.
        self.navigator_dock.hide()
        self.inspector_dock.hide()

        # Wire navigator file activation -> open file.
        self.navigator_dock.file_activated.connect(self._open_file)

        # View menu toggles use the dock's built-in toggle action.
        if hasattr(self, "_view_menu"):
            self._view_menu.addAction(self.navigator_dock.toggleViewAction())
            self._view_menu.addAction(self.inspector_dock.toggleViewAction())

    def _activate_next_tab(self):
        if not hasattr(self, "workspace_tabs"):
            return
        n = self.workspace_tabs.count()
        if n <= 1:
            return
        self.workspace_tabs.setCurrentIndex((self.workspace_tabs.currentIndex() + 1) % n)

    def _activate_prev_tab(self):
        if not hasattr(self, "workspace_tabs"):
            return
        n = self.workspace_tabs.count()
        if n <= 1:
            return
        self.workspace_tabs.setCurrentIndex((self.workspace_tabs.currentIndex() - 1) % n)

    def _close_active_tab(self):
        if not hasattr(self, "workspace_tabs"):
            return
        # Phase 2: only one persistent Analysis tab, can't be closed yet.
        # Phase 3 will allow closing additional session tabs.
        idx = self.workspace_tabs.currentIndex()
        if idx > 0:
            self.workspace_tabs.removeTab(idx)

    def _update_status_bar(self):
        if self.current_file_path:
            self.status_file_label.setText(os.path.basename(self.current_file_path))
            self.topbar_file_label.setText(os.path.basename(self.current_file_path))
        else:
            self.status_file_label.setText("No file")
            self.topbar_file_label.setText("No file loaded")
        self.status_rows_label.setText(f"{0 if self.df is None else len(self.df)} rows")
        self.status_col_label.setText(self.col_combo.currentText() or "No column" if hasattr(self, "col_combo") else "No column")
        self.status_alg_label.setText(self.alg_combo.currentText() if hasattr(self, "alg_combo") else "kmeans")

    def _select_page(self, index: int):
        self.stack.setCurrentIndex(index)
        for i, btn in enumerate(self.nav_buttons):
            btn.setChecked(i == index)
        titles = [
            ("Setup", "Choose your data and clustering parameters."),
            ("Cleaning", "Refine your text. Toggles preview live below."),
            ("Results", "Review clusters, edit names, visualize, and export."),
        ]
        title, subtitle = titles[index]
        self.page_title_label.setText(title)
        self.page_subtitle_label.setText(subtitle)

    # ------------------------------------------------------------------ #
    # Setup page                                                         #
    # ------------------------------------------------------------------ #

    def _build_setup_tab(self):
        layout = QtWidgets.QVBoxLayout(self.setup_page)
        layout.setSpacing(14)

        file_group = QtWidgets.QGroupBox("Data Source")
        file_group.setSizePolicy(
            QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Maximum
        )
        file_group_layout = QtWidgets.QVBoxLayout(file_group)
        file_group_layout.setContentsMargins(12, 14, 12, 12)
        self.data_source_panel = DataSourcePanel()
        self.data_source_panel.file_picked.connect(self._open_file)
        self.data_source_panel.browse_requested.connect(self.select_file)
        self.data_source_panel.change_requested.connect(self.select_file)
        self.data_source_panel.sheet_changed.connect(self._load_sheet)
        # Compatibility shims so existing code that touches these still works.
        self.file_label = QtWidgets.QLabel()
        self.file_label.hide()
        self.sheet_combo = self.data_source_panel._sheet_combo
        file_group_layout.addWidget(self.data_source_panel)
        layout.addWidget(file_group)

        params_group = QtWidgets.QGroupBox("Clustering Parameters")
        params_layout = QtWidgets.QGridLayout(params_group)
        params_layout.setHorizontalSpacing(12)
        params_layout.setVerticalSpacing(10)

        self.col_combo = QtWidgets.QComboBox()
        self.col_combo.setToolTip("Column containing the text to cluster")
        self.col_combo.currentTextChanged.connect(self.refresh_cleaning_preview)
        self.col_combo.currentTextChanged.connect(lambda *_: self._update_status_bar())
        self.col_combo.currentTextChanged.connect(
            lambda name: self.data_source_panel.set_highlight_column(name or None)
        )
        self.alg_combo = QtWidgets.QComboBox()
        # Order reflects recommended-first: kmeans is the safe default,
        # hdbscan auto-tunes density-based clusters, agglomerative is fine
        # on small data, dbscan is last because its eps parameter is the
        # #1 footgun on TF-IDF (cosine distances vary wildly per corpus).
        # If hdbscan is selected without the package installed, cluster_texts
        # raises ImportError with a clear install hint.
        # Display label vs engine value: "kmeans++" surfaces the init algorithm
        # actually in use; the engine still receives the bare "kmeans" string.
        # Use currentData() / self.current_algorithm() for the engine value;
        # currentText() for display purposes (status bar, persisted setting).
        for display, engine in (
            ("kmeans++", "kmeans"),
            ("hdbscan", "hdbscan"),
            ("agglomerative", "agglomerative"),
            ("dbscan", "dbscan"),
        ):
            self.alg_combo.addItem(display, userData=engine)
        self.alg_combo.setItemData(
            3,
            "DBSCAN requires manually tuning `eps`, which is brittle on TF-IDF "
            "data. For density-based clustering, prefer HDBSCAN — it auto-tunes "
            "and produces more stable clusters.",
            QtCore.Qt.ItemDataRole.ToolTipRole,
        )
        last_alg = self.settings.get("last_algorithm", "kmeans")
        self._set_algorithm(last_alg)
        self.alg_combo.currentTextChanged.connect(self._on_alg_change)
        self.alg_combo.currentTextChanged.connect(lambda *_: self._update_status_bar())
        self.k_spin = QtWidgets.QSpinBox()
        self.k_spin.setRange(1, 999)
        self.k_spin.setValue(5)
        self.k_spin.setToolTip("Number of clusters (kmeans / agglomerative)")
        self.name_top_spin = QtWidgets.QSpinBox()
        self.name_top_spin.setRange(1, 50)
        self.name_top_spin.setValue(3)
        self.name_top_spin.setToolTip("Top-N keywords used to auto-name each cluster")
        self.joiner_edit = QtWidgets.QLineEdit("_")
        self.joiner_edit.setToolTip("Character used to join keywords in cluster names")
        self.vis_combo = QtWidgets.QComboBox()
        self.vis_combo.addItems(["pca", "tsne"])
        self.out_edit = QtWidgets.QLineEdit()
        self.out_edit.setPlaceholderText("Output file path (.xlsx, .csv, or .json)")

        # Wrap n_clusters spin + Suggest (silhouette) and Elbow buttons so
        # users can call find_optimal_k without leaving the Setup tab.
        k_field = QtWidgets.QWidget()
        k_field_layout = QtWidgets.QHBoxLayout(k_field)
        k_field_layout.setContentsMargins(0, 0, 0, 0)
        k_field_layout.setSpacing(6)
        k_field_layout.addWidget(self.k_spin, 1)
        self.suggest_k_btn = QtWidgets.QPushButton("Suggest")
        self.suggest_k_btn.setToolTip("Recommend a cluster count via silhouette analysis")
        self.suggest_k_btn.setProperty("flat", "true")
        self.suggest_k_btn.setIcon(_icon("fa5s.magic"))
        self.suggest_k_btn.clicked.connect(lambda: self.suggest_optimal_k(method="silhouette"))
        k_field_layout.addWidget(self.suggest_k_btn)
        self.elbow_k_btn = QtWidgets.QPushButton("Elbow")
        self.elbow_k_btn.setToolTip("Recommend a cluster count via the elbow method (inertia bend)")
        self.elbow_k_btn.setProperty("flat", "true")
        self.elbow_k_btn.setIcon(_icon("fa5s.chart-line"))
        self.elbow_k_btn.clicked.connect(lambda: self.suggest_optimal_k(method="elbow"))
        k_field_layout.addWidget(self.elbow_k_btn)

        # Algorithm row: combo + Advanced TF-IDF + Compare Algorithms helpers.
        alg_field = QtWidgets.QWidget()
        alg_field_layout = QtWidgets.QHBoxLayout(alg_field)
        alg_field_layout.setContentsMargins(0, 0, 0, 0)
        alg_field_layout.setSpacing(6)
        alg_field_layout.addWidget(self.alg_combo, 1)
        self.advanced_tfidf_btn = QtWidgets.QPushButton("Advanced…")
        self.advanced_tfidf_btn.setToolTip("TF-IDF: min_df / max_df / n-grams / hashing")
        self.advanced_tfidf_btn.setProperty("flat", "true")
        self.advanced_tfidf_btn.clicked.connect(self.open_advanced_tfidf)
        alg_field_layout.addWidget(self.advanced_tfidf_btn)
        self.compare_algos_btn = QtWidgets.QPushButton("Compare")
        self.compare_algos_btn.setToolTip("Run all algorithms and compare metrics")
        self.compare_algos_btn.setProperty("flat", "true")
        self.compare_algos_btn.setIcon(_icon("fa5s.balance-scale"))
        self.compare_algos_btn.clicked.connect(self.compare_algorithms_dialog)
        alg_field_layout.addWidget(self.compare_algos_btn)

        widgets = [
            ("Text column:", self.col_combo),
            ("Algorithm:", alg_field),
            ("n_clusters:", k_field),
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
        params_layout.setColumnStretch(1, 1)
        params_layout.setColumnStretch(3, 1)
        layout.addWidget(params_group)
        layout.addStretch(1)

    # ------------------------------------------------------------------ #
    # Cleaning page                                                      #
    # ------------------------------------------------------------------ #

    def _build_cleaning_tab(self):
        layout = QtWidgets.QVBoxLayout(self.cleaning_page)
        layout.setSpacing(14)

        # Recipe row
        recipes = QtWidgets.QGroupBox("Cleaning Recipe")
        recipes_layout = QtWidgets.QHBoxLayout(recipes)
        recipes_layout.setSpacing(8)
        self.cleaning_recipe_combo = QtWidgets.QComboBox()
        self.cleaning_recipe_combo.setMinimumWidth(220)
        load_recipe_btn = QtWidgets.QPushButton("Load")
        load_recipe_btn.setIcon(_icon("fa5s.upload"))
        load_recipe_btn.clicked.connect(self.load_selected_cleaning_recipe)
        save_recipe_btn = QtWidgets.QPushButton("Save")
        save_recipe_btn.setIcon(_icon("fa5s.save"))
        save_recipe_btn.clicked.connect(self.save_current_cleaning_recipe)
        delete_recipe_btn = QtWidgets.QPushButton("Delete")
        delete_recipe_btn.setIcon(_icon("fa5s.trash-alt"))
        delete_recipe_btn.clicked.connect(self.delete_selected_cleaning_recipe)
        recipes_layout.addWidget(QtWidgets.QLabel("Recipe:"))
        recipes_layout.addWidget(self.cleaning_recipe_combo)
        recipes_layout.addWidget(load_recipe_btn)
        recipes_layout.addWidget(save_recipe_btn)
        recipes_layout.addWidget(delete_recipe_btn)
        recipes_layout.addStretch(1)
        refresh_preview_btn = QtWidgets.QPushButton("Refresh Preview")
        refresh_preview_btn.setIcon(_icon("fa5s.sync-alt"))
        refresh_preview_btn.clicked.connect(self.refresh_cleaning_preview)
        recipes_layout.addWidget(refresh_preview_btn)
        layout.addWidget(recipes)

        # Toggle cards in three columns
        self.replace_missing_check = QtWidgets.QCheckBox("Replace missing values")
        self.replace_missing_check.setChecked(True)
        self.missing_value_text_edit = QtWidgets.QLineEdit()
        self.missing_value_text_edit.setPlaceholderText("Replacement text (blank = drop)")
        self.trim_whitespace_check = QtWidgets.QCheckBox("Trim leading/trailing whitespace")
        self.trim_whitespace_check.setChecked(True)
        self.lowercase_cleaning_check = QtWidgets.QCheckBox("Convert to lowercase")
        self.lowercase_cleaning_check.setChecked(True)
        self.collapse_whitespace_check = QtWidgets.QCheckBox("Collapse repeated whitespace")
        self.collapse_whitespace_check.setChecked(True)
        self.remove_punctuation_check = QtWidgets.QCheckBox("Remove punctuation")
        self.remove_numbers_check = QtWidgets.QCheckBox("Remove standalone numbers")
        self.remove_urls_check = QtWidgets.QCheckBox("Remove URLs")
        self.remove_emails_check = QtWidgets.QCheckBox("Remove email addresses")
        self.dedupe_cleaned_rows_check = QtWidgets.QCheckBox("Drop duplicate cleaned rows")
        self.lemmatize_check = QtWidgets.QCheckBox("Lemmatize words (NLTK)")
        self.lemmatize_check.setToolTip(
            "Reduce words to their base form (e.g. 'running' → 'run'). "
            "Requires NLTK; corpora auto-download on first use."
        )
        self.regex_pattern_edit = QtWidgets.QLineEdit()
        self.regex_pattern_edit.setPlaceholderText("Pattern (Python regex)")
        self.regex_replacement_edit = QtWidgets.QLineEdit()
        self.regex_replacement_edit.setPlaceholderText("Replacement (use \\1 for groups)")

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
            self.lemmatize_check,
        ]
        for toggle in toggles:
            toggle.toggled.connect(self.refresh_cleaning_preview)
        self.regex_pattern_edit.editingFinished.connect(self.refresh_cleaning_preview)
        self.regex_replacement_edit.editingFinished.connect(self.refresh_cleaning_preview)
        self.missing_value_text_edit.editingFinished.connect(self.refresh_cleaning_preview)

        cards_row = QtWidgets.QHBoxLayout()
        cards_row.setSpacing(12)
        cards_row.addWidget(self._cleaning_card(
            "Whitespace & Case",
            "fa5s.text-height",
            [self.trim_whitespace_check, self.collapse_whitespace_check, self.lowercase_cleaning_check],
        ), 1)
        cards_row.addWidget(self._cleaning_card(
            "Tokens & Symbols",
            "fa5s.filter",
            [self.remove_punctuation_check, self.remove_numbers_check, self.dedupe_cleaned_rows_check, self.lemmatize_check],
        ), 1)
        cards_row.addWidget(self._cleaning_card(
            "Patterns",
            "fa5s.code",
            [self.remove_urls_check, self.remove_emails_check],
        ), 1)
        layout.addLayout(cards_row)

        # Missing values + regex group
        advanced = QtWidgets.QGroupBox("Custom Rules")
        advanced_layout = QtWidgets.QGridLayout(advanced)
        advanced_layout.setHorizontalSpacing(12)
        advanced_layout.setVerticalSpacing(8)
        advanced_layout.addWidget(self.replace_missing_check, 0, 0)
        advanced_layout.addWidget(self.missing_value_text_edit, 0, 1, 1, 3)
        advanced_layout.addWidget(QtWidgets.QLabel("Regex find:"), 1, 0)
        advanced_layout.addWidget(self.regex_pattern_edit, 1, 1)
        advanced_layout.addWidget(QtWidgets.QLabel("Replace with:"), 1, 2)
        advanced_layout.addWidget(self.regex_replacement_edit, 1, 3)
        advanced_layout.setColumnStretch(1, 2)
        advanced_layout.setColumnStretch(3, 2)
        layout.addWidget(advanced)

        # Preview
        preview_group = QtWidgets.QGroupBox("Live Preview")
        preview_layout = QtWidgets.QVBoxLayout(preview_group)
        metrics_layout = QtWidgets.QGridLayout()
        metrics_layout.setHorizontalSpacing(20)
        metrics_layout.setVerticalSpacing(4)
        self.cleaning_source_label = QtWidgets.QLabel("No source column")
        self.cleaning_output_label = QtWidgets.QLabel("No cleaned column yet")
        self.cleaning_rows_label = QtWidgets.QLabel("0")
        self.cleaning_cluster_rows_label = QtWidgets.QLabel("0")
        self.cleaning_empty_rows_label = QtWidgets.QLabel("0")
        self.cleaning_deduped_rows_label = QtWidgets.QLabel("0")
        metrics = [
            ("Source column", self.cleaning_source_label),
            ("Cleaned column", self.cleaning_output_label),
            ("Rows before", self.cleaning_rows_label),
            ("Rows for clustering", self.cleaning_cluster_rows_label),
            ("Empty after cleaning", self.cleaning_empty_rows_label),
            ("Deduped rows", self.cleaning_deduped_rows_label),
        ]
        for index, (label_text, widget) in enumerate(metrics):
            cap = QtWidgets.QLabel(label_text)
            cap.setProperty("role", "subtle")
            metrics_layout.addWidget(cap, (index // 3) * 2, index % 3)
            metrics_layout.addWidget(widget, (index // 3) * 2 + 1, index % 3)
        preview_layout.addLayout(metrics_layout)

        self.cleaning_preview_table = QtWidgets.QTableWidget(0, 2)
        self.cleaning_preview_table.setHorizontalHeaderLabels(["Original", "Cleaned"])
        self.cleaning_preview_table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.cleaning_preview_table.setAlternatingRowColors(True)
        self.cleaning_preview_table.verticalHeader().setVisible(False)
        preview_layout.addWidget(self.cleaning_preview_table, 1)
        layout.addWidget(preview_group, 1)

    def _cleaning_card(self, title: str, icon_name: str, toggles: list) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("Card")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(8)
        header = QtWidgets.QHBoxLayout()
        header.setSpacing(8)
        if _HAS_QTA:
            icon_label = QtWidgets.QLabel()
            icon_label.setPixmap(_icon(icon_name).pixmap(16, 16))
            header.addWidget(icon_label)
        title_label = QtWidgets.QLabel(title)
        title_label.setProperty("role", "h2")
        title_label.setStyleSheet("font-size: 14px;")
        header.addWidget(title_label)
        header.addStretch(1)
        layout.addLayout(header)
        for toggle in toggles:
            layout.addWidget(toggle)
        layout.addStretch(1)
        return card

    # ------------------------------------------------------------------ #
    # Results page                                                       #
    # ------------------------------------------------------------------ #

    def _build_results_tab(self):
        layout = QtWidgets.QVBoxLayout(self.results_page)
        layout.setSpacing(14)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        layout.addWidget(splitter, 1)

        # Left: cluster cards
        names_group = QtWidgets.QGroupBox("Clusters")
        names_layout = QtWidgets.QVBoxLayout(names_group)
        names_help = QtWidgets.QLabel("Edit names below. Click Save Results when ready.")
        names_help.setProperty("role", "muted")
        names_layout.addWidget(names_help)
        self.names_scroll = QtWidgets.QScrollArea()
        self.names_scroll.setWidgetResizable(True)
        self.names_scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.names_container = QtWidgets.QWidget()
        self.names_form = QtWidgets.QVBoxLayout(self.names_container)
        self.names_form.setSpacing(8)
        self.names_form.addStretch(1)
        self.names_scroll.setWidget(self.names_container)
        names_layout.addWidget(self.names_scroll, 1)
        splitter.addWidget(names_group)

        # Right: viz + log
        right = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(12)

        viz_group = QtWidgets.QGroupBox("Cluster Visualization")
        viz_layout = QtWidgets.QVBoxLayout(viz_group)
        viz_top = QtWidgets.QHBoxLayout()
        viz_top.setSpacing(8)
        viz_method_label = QtWidgets.QLabel("Method:")
        viz_top.addWidget(viz_method_label)
        viz_top.addWidget(self._inline_label("Set on Setup tab via 'Visualization'"))
        viz_top.addStretch(1)
        self.viz_btn = QtWidgets.QPushButton("Visualize")
        self.viz_btn.setIcon(_icon("fa5s.chart-scatter"))
        self.viz_btn.setEnabled(False)
        self.viz_btn.clicked.connect(self.visualize_clusters)
        viz_top.addWidget(self.viz_btn)
        viz_layout.addLayout(viz_top)
        self.embedding_canvas = EmbeddingCanvas()
        viz_layout.addWidget(self.embedding_canvas, 1)
        right_layout.addWidget(viz_group, 2)

        log_group = QtWidgets.QGroupBox("Status Log")
        log_layout = QtWidgets.QVBoxLayout(log_group)
        log_top = QtWidgets.QHBoxLayout()
        log_top.addStretch(1)
        copy_log_btn = QtWidgets.QPushButton("Copy")
        copy_log_btn.setIcon(_icon("fa5s.copy"))
        copy_log_btn.setProperty("flat", "true")
        copy_log_btn.clicked.connect(self._copy_log)
        clear_log_btn = QtWidgets.QPushButton("Clear")
        clear_log_btn.setIcon(_icon("fa5s.trash-alt"))
        clear_log_btn.setProperty("flat", "true")
        clear_log_btn.clicked.connect(self.clear_log)
        log_top.addWidget(copy_log_btn)
        log_top.addWidget(clear_log_btn)
        log_layout.addLayout(log_top)
        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(120)
        log_layout.addWidget(self.log, 1)
        right_layout.addWidget(log_group, 1)

        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([380, 720])

    def _inline_label(self, text: str) -> QtWidgets.QLabel:
        label = QtWidgets.QLabel(text)
        label.setProperty("role", "subtle")
        return label

    def _copy_log(self):
        QtWidgets.QApplication.clipboard().setText(self.log.toPlainText())
        show_toast(self, "Log copied to clipboard", level="success")

    # ------------------------------------------------------------------ #
    # Cleaning recipe & config helpers                                   #
    # ------------------------------------------------------------------ #

    def current_algorithm(self) -> str:
        """Engine-side algorithm name for the current dropdown selection.

        Display labels (e.g. ``kmeans++``) and engine values (e.g. ``kmeans``)
        differ; the engine name is stored as Qt itemData and read via
        currentData(). Falls back to currentText() for older builds where
        items were added without itemData.
        """
        data = self.alg_combo.currentData()
        return str(data) if data is not None else self.alg_combo.currentText()

    def _set_algorithm(self, engine_value: str) -> None:
        """Select the dropdown item whose engine name (itemData) matches.

        Used both for restoring the last-used algorithm at startup and for
        the Compare Algorithms dialog's "Use {best}" button.
        """
        for i in range(self.alg_combo.count()):
            if self.alg_combo.itemData(i) == engine_value:
                self.alg_combo.setCurrentIndex(i)
                return
        # Fall back to text match for forward-compat / unknown values.
        idx = self.alg_combo.findText(engine_value)
        if idx >= 0:
            self.alg_combo.setCurrentIndex(idx)

    def _on_alg_change(self, *_args):
        self.k_spin.setEnabled(self.current_algorithm() != "dbscan")

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
            lemmatize=self.lemmatize_check.isChecked(),
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
            "lemmatize": config.lemmatize,
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
        self.lemmatize_check.setChecked(bool(recipe.get("lemmatize", False)))

    def _refresh_cleaning_recipe_combo(self):
        current = self.cleaning_recipe_combo.currentText()
        self.cleaning_recipe_combo.blockSignals(True)
        self.cleaning_recipe_combo.clear()
        self.cleaning_recipe_combo.addItems(sorted(self.cleaning_recipes.keys()))
        if current in self.cleaning_recipes:
            self.cleaning_recipe_combo.setCurrentText(current)
        self.cleaning_recipe_combo.blockSignals(False)

    def _persist_recipes(self):
        self.settings["cleaning_recipes"] = dict(self.cleaning_recipes)
        app_settings.save(self.settings)

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
        self._persist_recipes()
        self.log_msg(f"âœ“ Cleaning recipe saved: {clean_name}")
        show_toast(self, f"Recipe '{clean_name}' saved", level="success")

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
        self.log_msg(f"âœ“ Cleaning recipe loaded: {recipe_name}")
        show_toast(self, f"Recipe '{recipe_name}' loaded", level="info")

    def delete_selected_cleaning_recipe(self):
        recipe_name = self.cleaning_recipe_combo.currentText().strip()
        if not recipe_name:
            show_warning(self, "No Recipe", "Select a cleaning recipe to delete.")
            return
        if QtWidgets.QMessageBox.question(self, "Delete Recipe", f"Delete recipe '{recipe_name}'?") != QtWidgets.QMessageBox.Yes:
            return
        self.cleaning_recipes.pop(recipe_name, None)
        self._refresh_cleaning_recipe_combo()
        self._persist_recipes()
        self.log_msg(f"âœ“ Cleaning recipe deleted: {recipe_name}")

    def _get_cleaned_column_name(self, source_column):
        return f"{source_column}_cleaned"

    def refresh_cleaning_preview(self):
        if not hasattr(self, "cleaning_preview_table"):
            return
        self.cleaning_preview_table.setRowCount(0)
        if self.df is None or not self.col_combo.currentText():
            self.cleaning_source_label.setText("No source column")
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
            self.log_msg(f"âœ— Cleaning preview failed: {error}")

    # ------------------------------------------------------------------ #
    # Status / log                                                       #
    # ------------------------------------------------------------------ #

    def status_message(self, message):
        self.statusBar().showMessage(message, 4000)

    def log_msg(self, msg):
        formatted = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}"
        self.log.appendPlainText(formatted)
        # Mirror to diagnostics window if open.
        if self.diagnostics_window is not None and self.diagnostics_window.isVisible():
            self.diagnostics_window.append_entry(msg)

    def clear_log(self):
        self.log.clear()

    def show_imprint(self):
        from textanalyzer.ui.about import AboutDialog
        dialog = AboutDialog(self.owner_name, self.owner_contact, self.owner_website, self.owner_bmc, self)
        dialog.exec()

    def show_settings_window(self):
        if self.settings_window is None:
            from textanalyzer.ui.settings_window import SettingsWindow
            self.settings_window = SettingsWindow(self.settings, self)
            self.settings_window.settings_changed.connect(self._on_settings_changed)
        self.settings_window.show()
        self.settings_window.raise_()
        self.settings_window.activateWindow()

    def show_diagnostics_window(self):
        if self.diagnostics_window is None:
            from textanalyzer.ui.diagnostics_window import DiagnosticsWindow
            self.diagnostics_window = DiagnosticsWindow(self)
            # Seed with existing in-page log content.
            existing = self.log.toPlainText().splitlines() if hasattr(self, "log") else []
            self.diagnostics_window.set_history(existing)
        self.diagnostics_window.show()
        self.diagnostics_window.raise_()
        self.diagnostics_window.activateWindow()

    def _on_settings_changed(self, change: dict) -> None:
        # Theme changes go through ThemeManager so all windows update.
        if "theme" in change and self.theme_manager is not None:
            target = change["theme"]
            if self.theme_manager.mode != target:
                self.theme_manager.apply(target)
                self.settings["theme"] = target
                app_settings.save(self.settings)
        if "recent_files" in change:
            self._refresh_recent_files_menu()

    def toggle_theme(self):
        if self.theme_manager is None:
            return
        new_mode = self.theme_manager.toggle()
        self.settings["theme"] = new_mode
        app_settings.save(self.settings)
        if self.wordcloud_builder is not None:
            self.wordcloud_builder.update()
        if self.settings_window is not None:
            self.settings_window.sync_theme_from_external(new_mode)
        show_toast(self, f"Theme: {new_mode.title()}", level="info")

    # ------------------------------------------------------------------ #
    # File / sheet loading                                               #
    # ------------------------------------------------------------------ #

    def current_column_name(self):
        return self.col_combo.currentText().strip()

    def current_sheet_name(self):
        return self.sheet_combo.currentText().strip()

    def select_file(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Select Input File",
            "",
            "Supported files (*.xlsx *.xlsm *.xltx *.xltm *.xls *.xlsb *.ods *.csv *.json)"
            ";;Excel files (*.xlsx *.xlsm *.xltx *.xltm *.xls *.xlsb)"
            ";;OpenDocument Spreadsheet (*.ods)"
            ";;CSV files (*.csv)"
            ";;JSON files (*.json)",
        )
        if not path:
            return
        self._open_file(path)

    def _open_file(self, path: str):
        if not os.path.exists(path):
            show_error(self, "File not found", f"Could not find: {path}")
            return
        self.current_file_path = path
        self.file_label.setText(path)
        self.setWindowTitle(f"{self.app_title} {self.app_version} — {os.path.basename(path)}")
        base, ext = os.path.splitext(path)
        output_ext = ext if ext in {".xlsx", ".csv", ".json"} else ".xlsx"
        self.out_edit.setText(base + "_clustered" + output_ext)
        try:
            sheet_names = get_sheet_names(path)
            current = sheet_names[0] if sheet_names else ""
            self.data_source_panel.show_loaded_state(
                file_path=path, sheets=sheet_names, current_sheet=current
            )
            if sheet_names:
                self._load_sheet(current)
            ext_label = get_file_extension(path)
            if ext_label in {".csv", ".json"}:
                self.log_msg(f"✓ Loaded {ext_label[1:].upper()} source with 1 table: {sheet_names[0]}")
            else:
                self.log_msg(f"✓ Found {len(sheet_names)} sheet(s): {', '.join(sheet_names)}")
            app_settings.add_recent_file(self.settings, path)
            app_settings.save(self.settings)
            self._refresh_recent_files_menu()
        except Exception as error:
            show_error(self, "Error", f"Failed to read file structure: {error}")
            self.log_msg(f"✗ Error reading file structure: {error}")
            self.sidebar_wordcloud_btn.setEnabled(False)

    def _refresh_recent_files_menu(self):
        if not hasattr(self, "recent_menu"):
            return
        self.recent_menu.clear()
        recents = self.settings.get("recent_files", []) or []
        recents = [p for p in recents if os.path.exists(p)]
        # Mirror into the navigator dock if it exists.
        if hasattr(self, "navigator_dock"):
            self.navigator_dock.set_recent_files(recents)
        if not recents:
            empty = QtGui.QAction("(empty)", self)
            empty.setEnabled(False)
            self.recent_menu.addAction(empty)
            return
        for p in recents:
            action = QtGui.QAction(os.path.basename(p), self)
            action.setToolTip(p)
            action.triggered.connect(lambda _checked=False, target=p: self._open_file(target))
            self.recent_menu.addAction(action)
        self.recent_menu.addSeparator()
        clear_action = QtGui.QAction("Clear Recent", self)
        clear_action.triggered.connect(self._clear_recent_files)
        self.recent_menu.addAction(clear_action)

    def _clear_recent_files(self):
        self.settings["recent_files"] = []
        app_settings.save(self.settings)
        self._refresh_recent_files_menu()

    def _load_sheet(self, sheet_name):
        if not self.current_file_path or not sheet_name:
            return
        try:
            ext = get_file_extension(self.current_file_path)
            from cluster_tool import EXCEL_INPUT_EXTENSIONS
            selected_sheet = sheet_name if ext in EXCEL_INPUT_EXTENSIONS else None
            self.df = load_table(self.current_file_path, sheet_name=selected_sheet)
            cols = list(self.df.columns)
            self.col_combo.blockSignals(True)
            self.col_combo.clear()
            self.col_combo.addItems(cols)
            self.col_combo.blockSignals(False)
            preferred = self.settings.get("last_column", "")
            if preferred and preferred in cols:
                self.col_combo.setCurrentText(preferred)
            elif cols:
                self.col_combo.setCurrentText(cols[0])
            if cols:
                self.sidebar_wordcloud_btn.setEnabled(True)
            else:
                self.sidebar_wordcloud_btn.setEnabled(False)
            source_label = "sheet" if ext in EXCEL_INPUT_EXTENSIONS else "table"
            file_size_kb = os.path.getsize(self.current_file_path) / 1024
            self.log_msg(f"✓ Loaded {source_label} '{sheet_name}': {len(self.df)} rows, {len(cols)} columns, {file_size_kb:.1f} KB")
            self.log_msg(f"  Columns: {', '.join(cols)}")
            self.data_source_panel.update_preview(
                self.df, highlight_column=self.col_combo.currentText() or None
            )
            self.refresh_cleaning_preview()
            self._update_status_bar()
            if self.wordcloud_builder is not None:
                self.wordcloud_builder.refresh_from_app()
        except Exception as error:
            show_error(self, "Error", f"Failed to load data: {error}")
            self.log_msg(f"✗ Error loading data: {error}")

    # ------------------------------------------------------------------ #
    # Clustering (real worker)                                           #
    # ------------------------------------------------------------------ #

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
        if self.controller._cluster_thread is not None and self.controller._cluster_thread.isRunning():
            show_warning(self, "Already running", "A clustering job is already in progress.")
            return
        try:
            n_clusters = int(self.k_spin.value())
            if self.current_algorithm() != "dbscan" and n_clusters < 2:
                show_warning(self, "Invalid parameter", "n_clusters must be at least 2")
                return
            top_n = int(self.name_top_spin.value())
        except ValueError as error:
            show_error(self, "Invalid input", f"Please enter valid numbers: {error}")
            return

        # Sync session df to controller before running.
        self.controller.session.df = self.df

        config = CleaningConfigModel(
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
            lemmatize=self.lemmatize_check.isChecked(),
        )
        self.controller.run_clustering(
            column=col,
            config=config,
            algorithm=self.current_algorithm(),
            n_clusters=n_clusters,
            top_n=top_n,
            joiner=self.joiner_edit.text(),
            vectorize_kwargs={
                "max_features": self._tfidf_settings["max_features"],
                "min_df": self._tfidf_settings["min_df"],
                "max_df": self._tfidf_settings["max_df"],
                "ngram_range": self._tfidf_settings["ngram_range"],
                "use_hashing": self._tfidf_settings["use_hashing"],
            },
        )

    def cancel_clustering(self):
        self.controller.cancel_clustering()

    def _set_running_state(self, running: bool):
        self.run_btn.setEnabled(not running)
        self.cancel_btn.setVisible(running)
        self.progress.setVisible(running)
        self.progress_label.setVisible(running)
        if running:
            self.progress.setValue(0)
            self.progress_label.setText("Starting\u2026")
            self.save_btn.setEnabled(False)
            self.viz_btn.setEnabled(False)
            self.save_model_btn.setEnabled(False)

    @QtCore.Slot(int, str)
    def _on_cluster_progress(self, value: int, message: str):
        self.progress.setValue(value)
        self.progress_label.setText(message)

    def _on_ctrl_cluster_finished(self, model) -> None:
        """Update view from :class:`ClusterResultModel` produced by controller."""
        try:
            self.latest_cleaning_result = self.controller.session.cleaning_result
            # Keep legacy instance vars in sync for backwards compat.
            self.X = model.X
            self.vectorizer = model.vectorizer
            self.model = model.model
            self.labels = model.labels
            self._kept_labels = model.kept_labels
            self.top_keywords = model.top_keywords
            self.cluster_names = model.cluster_names

            cleaned_col = model.cleaned_column_name
            if cleaned_col and cleaned_col not in [self.col_combo.itemText(i) for i in range(self.col_combo.count())]:
                self.col_combo.addItem(cleaned_col)
            self.refresh_cleaning_preview()

            self.populate_name_entries()
            self._select_page(self.PAGE_RESULTS)

            self.save_btn.setEnabled(True)
            self.viz_btn.setEnabled(True)
            self.save_model_btn.setEnabled(True)
            n_input = len(self.controller.session.cleaning_result.cluster_input_texts) if self.controller.session.cleaning_result else 0
            show_toast(self, f"Clustered {n_input} rows into {model.n_clusters} clusters", level="success")
            self.settings["last_algorithm"] = self.current_algorithm()
            app_settings.save(self.settings)
        finally:
            self._set_running_state(False)

    def _on_ctrl_cluster_failed(self, error_message: str) -> None:
        show_error(self, "Error during clustering", error_message)
        self._set_running_state(False)

    def populate_name_entries(self):
        # clear existing entries
        while self.names_form.count():
            item = self.names_form.takeAt(0)
            if item is None:
                continue
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        self.name_entries = {}
        for cid in sorted(self.cluster_names.keys()):
            card = self._build_cluster_card(cid)
            self.names_form.addWidget(card)
        self.names_form.addStretch(1)

    def _build_cluster_card(self, cid: int) -> QtWidgets.QFrame:
        card = QtWidgets.QFrame()
        card.setObjectName("ClusterCard")
        layout = QtWidgets.QVBoxLayout(card)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(6)

        header = QtWidgets.QHBoxLayout()
        header.setSpacing(8)
        chip = QtWidgets.QLabel(f"#{cid}")
        chip.setObjectName("ClusterChip")
        chip.setMaximumHeight(20)
        header.addWidget(chip)
        entry = QtWidgets.QLineEdit(self.cluster_names[cid])
        entry.setPlaceholderText("Cluster name")
        header.addWidget(entry, 1)
        # row count for this cluster
        if self.labels is not None:
            count = int(np.sum(self.labels == cid))
            count_label = QtWidgets.QLabel(f"{count} rows")
            count_label.setProperty("role", "muted")
            header.addWidget(count_label)
        layout.addLayout(header)

        keywords = [term for term, _score in self.top_keywords.get(cid, [])]
        if keywords:
            kw_row = QtWidgets.QHBoxLayout()
            kw_row.setSpacing(4)
            for term in keywords[:8]:
                pill = QtWidgets.QLabel(term)
                pill.setObjectName("KeywordPill")
                kw_row.addWidget(pill)
            kw_row.addStretch(1)
            layout.addLayout(kw_row)

        self.name_entries[cid] = entry
        return card

    # ------------------------------------------------------------------ #
    # Visualization                                                      #
    # ------------------------------------------------------------------ #

    def visualize_clusters(self):
        if self.X is None or self.labels is None:
            show_warning(self, "Nothing to visualize", "Run clustering first")
            return
        method = self.vis_combo.currentText()
        self.log_msg(f"Generating {method.upper()} visualization\u2026")
        if not self.embedding_canvas.is_available():
            try:
                visualize_embeddings(self.X, self.labels, method=method)
                self.log_msg(f"\u2713 {method.upper()} visualization displayed (matplotlib popup)")
            except Exception as error:
                self.log_msg(f"\u2717 Visualization failed: {error}")
                show_error(self, "Visualization failed", error)
            return
        self.viz_btn.setEnabled(False)
        self.controller.visualize(method)
        self._select_page(self.PAGE_RESULTS)

    def _on_ctrl_embedding_ready(self, embedding_2d, kept_labels, cluster_names) -> None:
        try:
            tokens = self.theme_manager.tokens() if self.theme_manager is not None else None
            bg = (tokens or {}).get("surface", "#ffffff")
            fg = (tokens or {}).get("text", "#1d2230")
            muted = (tokens or {}).get("text_muted", "#5b6271")
            method = self.vis_combo.currentText() if hasattr(self, 'vis_combo') else "pca"
            self.embedding_canvas.render_embedding(embedding_2d, kept_labels, method, bg, fg, muted)
            self.log_msg(f"\u2713 {method.upper()} visualization rendered")
        finally:
            self.viz_btn.setEnabled(True)

    # ------------------------------------------------------------------ #
    # Wordcloud / save / model                                           #
    # ------------------------------------------------------------------ #

    def open_wordcloud_builder(self):
        if self.df is None:
            show_warning(self, "No file", "Please select a file and table first")
            return
        col = self.current_column_name()
        if not col or col not in self.df.columns:
            show_warning(self, "No column", "Please select a text column first")
            return
        texts = coerce_text_column(self.df[col]).tolist()
        if self.wordcloud_builder is not None and self.wordcloud_builder.isVisible():
            self.wordcloud_builder.close()
        from textanalyzer.ui.wordcloud_window import WordCloudDialog
        self.wordcloud_builder = WordCloudDialog(self, texts, col, dataframe=self.df)
        self.wordcloud_builder.showMaximized()

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
            self.log_msg(f"âœ“ Results saved to {saved_path}")
            show_toast(self, f"Saved to {os.path.basename(saved_path)}", level="success")
        except Exception as error:
            self.log_msg(f"âœ— Save failed: {error}")
            show_error(self, "Save failed", error)

    def save_model(self):
        if self.model is None or self.vectorizer is None:
            show_warning(self, "Nothing to save", "Run clustering first")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Clustering Model", "", "Joblib files (*.joblib)")
        if not path:
            return
        try:
            IOService.save_model(self.model, self.vectorizer, self.cluster_names, self.top_keywords, path)
            self.log_msg(f"\u2713 Model saved to {path}")
            show_toast(self, f"Model saved: {os.path.basename(path)}", level="success")
        except Exception as error:
            self.log_msg(f"\u2717 Model save failed: {error}")
            show_error(self, "Save Failed", f"Failed to save model: {error}")

    def load_model_apply(self):
        """Load a saved .joblib model and apply it to the currently loaded data."""
        if self.df is None:
            show_warning(self, "No data", "Open a file first, then load a model to apply to it.")
            return
        col = self.current_column_name()
        if not col or col not in self.df.columns:
            show_warning(self, "No column", "Select a text column on the Setup tab first.")
            return
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Load Clustering Model", "", "Joblib files (*.joblib);;All files (*)"
        )
        if not path:
            return
        try:
            payload = IOService.load_model(path)
            loaded_model = payload["model"]
            loaded_vectorizer = payload["vectorizer"]
            loaded_cluster_names = payload.get("cluster_names", {}) or {}
            loaded_top_keywords = payload.get("top_keywords", {}) or {}

            texts = coerce_text_column(self.df[col]).tolist()
            X_new = loaded_vectorizer.transform(texts)
            if not hasattr(loaded_model, "predict"):
                raise RuntimeError(
                    "Loaded model has no .predict(). Save with the latest version "
                    "to enable prediction for non-KMeans algorithms."
                )
            labels = loaded_model.predict(X_new)
        except Exception as error:
            self.log_msg(f"\u2717 Load model failed: {error}")
            show_error(self, "Load Failed", f"Failed to load/apply model: {error}")
            return

        # Adopt the loaded artifacts as the current run state so Save Results,
        # Visualize, etc. work as if we had just trained.
        self.model = loaded_model
        self.vectorizer = loaded_vectorizer
        self.cluster_names = loaded_cluster_names
        self.top_keywords = loaded_top_keywords
        self.X = X_new
        self.labels = np.asarray(labels)
        self.df["cluster_label"] = self.labels
        self.df["cluster_name"] = [loaded_cluster_names.get(int(l), "") for l in self.labels]

        n_clusters = int(len({int(l) for l in self.labels if l != -1}))
        self.log_msg(
            f"\u2713 Loaded model from {os.path.basename(path)} \u2014 "
            f"{n_clusters} clusters applied to {len(self.df)} rows"
        )
        show_toast(self, f"Model applied: {n_clusters} clusters", level="success")
        if hasattr(self, "populate_name_entries"):
            try:
                self.populate_name_entries()
            except Exception:
                pass
        if hasattr(self, "save_btn"):
            self.save_btn.setEnabled(True)
        if hasattr(self, "viz_btn"):
            self.viz_btn.setEnabled(True)
        if hasattr(self, "save_model_btn"):
            self.save_model_btn.setEnabled(True)
        self._select_page(self.PAGE_RESULTS)
        self._update_status_bar()

    # ------------------------------------------------------------------ #
    # Advanced clustering helpers (Suggest K, Advanced TF-IDF, Compare)  #
    # ------------------------------------------------------------------ #

    def _prepare_setup_corpus(self):
        """Clean + vectorize current text column for ad-hoc analysis.

        Used by Suggest K and Compare Algorithms. Returns ``(vectorizer, X)``
        or ``None`` on failure (warnings already shown to the user).
        """
        if self.df is None:
            show_warning(self, "No data", "Open a file first.")
            return None
        col = self.current_column_name()
        if not col or col not in self.df.columns:
            show_warning(self, "No column", "Pick a text column on the Setup tab first.")
            return None
        try:
            cleaning_result = prepare_text_cleaning(
                self.df[col].tolist(), self._build_cleaning_config()
            )
        except Exception as exc:
            show_error(self, "Cleaning failed", f"Could not clean text: {exc}")
            return None
        if not cleaning_result.cluster_input_texts:
            show_warning(
                self,
                "Nothing to analyze",
                "Cleaning produced no usable rows. Adjust cleaning settings and retry.",
            )
            return None
        try:
            from cluster_tool import vectorize_texts as _vectorize
            vectorizer, X = _vectorize(
                cleaning_result.cluster_input_texts,
                max_features=self._tfidf_settings["max_features"],
                min_df=self._tfidf_settings["min_df"],
                max_df=self._tfidf_settings["max_df"],
                ngram_range=self._tfidf_settings["ngram_range"],
                use_hashing=self._tfidf_settings["use_hashing"],
            )
        except Exception as exc:
            show_error(self, "Vectorization failed", str(exc))
            return None
        return vectorizer, X

    def suggest_optimal_k(self, method: str = "silhouette"):
        """Run k-search via silhouette or elbow and offer to apply the result.

        method: 'silhouette' picks the k with the highest silhouette score
                (best for clusters that are visibly separable).
                'elbow' picks the k where inertia bends most sharply
                (better when silhouette is uniformly low).
        """
        from cluster_tool import find_optimal_k as _find_k

        method = method if method in ("silhouette", "elbow") else "silhouette"
        prepared = self._prepare_setup_corpus()
        if prepared is None:
            return
        _, X = prepared
        QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        self.log_msg(f"Searching for optimal k via {method} method…")
        try:
            result = _find_k(X, k_range=(2, min(15, X.shape[0] - 1)), method=method)
        except Exception as exc:
            QtWidgets.QApplication.restoreOverrideCursor()
            show_error(self, "Suggest K failed", str(exc))
            return
        QtWidgets.QApplication.restoreOverrideCursor()

        if method == "silhouette":
            scores = result.get("scores", {})
            detail_lines = [f"  k={k}: silhouette {s:.3f}" for k, s in sorted(scores.items())]
            detail_label = "Scores"
        else:
            inertias = result.get("inertias", {})
            detail_lines = [f"  k={k}: inertia {v:,.1f}" for k, v in sorted(inertias.items())]
            detail_label = "Inertias"
        detail_text = "\n".join(detail_lines) if detail_lines else "  (no values collected)"

        message = (
            f"{result['recommendation']}\n\n"
            f"Confidence: {result['confidence']}\n\n"
            f"{detail_label}:\n{detail_text}\n\n"
            f"Apply k = {result['optimal_k']} to the n_clusters field?"
        )
        title = "Suggested cluster count" if method == "silhouette" else "Elbow-method cluster count"
        button = QtWidgets.QMessageBox.question(
            self,
            title,
            message,
            QtWidgets.QMessageBox.Apply | QtWidgets.QMessageBox.Cancel,
            QtWidgets.QMessageBox.Apply,
        )
        if button == QtWidgets.QMessageBox.Apply:
            self.k_spin.setValue(int(result["optimal_k"]))
            self.log_msg(f"✓ Applied {method} k = {result['optimal_k']}")
        else:
            self.log_msg(f"{method.capitalize()} k = {result['optimal_k']} (not applied)")

    def open_advanced_tfidf(self):
        """Modal exposing min_df / max_df / ngram_range / hashing / max_features."""
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Advanced TF-IDF settings")
        dlg.setModal(True)
        form = QtWidgets.QFormLayout(dlg)
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(10)

        max_features_spin = QtWidgets.QSpinBox()
        max_features_spin.setRange(100, 1_000_000)
        max_features_spin.setSingleStep(500)
        max_features_spin.setValue(int(self._tfidf_settings["max_features"]))
        max_features_spin.setToolTip("Vocabulary cap for TF-IDF (top-N most common terms)")
        form.addRow("Max features", max_features_spin)

        min_df_spin = QtWidgets.QSpinBox()
        min_df_spin.setRange(1, 1000)
        min_df_spin.setValue(int(self._tfidf_settings["min_df"]))
        min_df_spin.setToolTip("Ignore terms appearing in fewer than this many documents")
        form.addRow("min_df (docs)", min_df_spin)

        max_df_spin = QtWidgets.QDoubleSpinBox()
        max_df_spin.setRange(0.05, 1.0)
        max_df_spin.setSingleStep(0.05)
        max_df_spin.setDecimals(2)
        max_df_spin.setValue(float(self._tfidf_settings["max_df"]))
        max_df_spin.setToolTip("Ignore terms appearing in more than this fraction of documents")
        form.addRow("max_df (fraction)", max_df_spin)

        ngram_combo = QtWidgets.QComboBox()
        ngram_combo.addItems(["unigrams (1,1)", "uni+bigrams (1,2)", "uni+bi+trigrams (1,3)"])
        ngram_map = {(1, 1): 0, (1, 2): 1, (1, 3): 2}
        ngram_combo.setCurrentIndex(ngram_map.get(tuple(self._tfidf_settings["ngram_range"]), 0))
        form.addRow("n-grams", ngram_combo)

        hashing_check = QtWidgets.QCheckBox("Use HashingVectorizer (memory-efficient; no keywords)")
        hashing_check.setChecked(bool(self._tfidf_settings["use_hashing"]))
        form.addRow("", hashing_check)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        form.addRow(buttons)

        if dlg.exec() != QtWidgets.QDialog.Accepted:
            return
        ngram_options = [(1, 1), (1, 2), (1, 3)]
        self._tfidf_settings.update({
            "max_features": int(max_features_spin.value()),
            "min_df": int(min_df_spin.value()),
            "max_df": float(max_df_spin.value()),
            "ngram_range": ngram_options[ngram_combo.currentIndex()],
            "use_hashing": bool(hashing_check.isChecked()),
        })
        self.log_msg(
            f"TF-IDF: max_features={self._tfidf_settings['max_features']}, "
            f"min_df={self._tfidf_settings['min_df']}, max_df={self._tfidf_settings['max_df']}, "
            f"ngram={self._tfidf_settings['ngram_range']}, "
            f"hashing={self._tfidf_settings['use_hashing']}"
        )

    def compare_algorithms_dialog(self):
        """Run KMeans / DBSCAN / Agglomerative on the same matrix and show metrics."""
        from cluster_tool import compare_algorithms as _compare, get_best_algorithm as _best

        prepared = self._prepare_setup_corpus()
        if prepared is None:
            return
        _, X = prepared

        progress = QtWidgets.QProgressDialog(
            "Running algorithms…", "Cancel", 0, 100, self
        )
        progress.setWindowTitle("Compare algorithms")
        progress.setWindowModality(QtCore.Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.setValue(0)

        cancelled = {"flag": False}
        progress.canceled.connect(lambda: cancelled.update(flag=True))

        def cb(stage: str, frac: float):
            if cancelled["flag"]:
                return
            progress.setLabelText(stage)
            progress.setValue(int(frac * 100))
            QtWidgets.QApplication.processEvents()

        n_clusters = self.k_spin.value()
        try:
            results = _compare(X, n_clusters=n_clusters, progress_callback=cb)
        except Exception as exc:
            progress.close()
            show_error(self, "Compare failed", str(exc))
            return
        progress.setValue(100)
        progress.close()
        if cancelled["flag"]:
            self.log_msg("Compare cancelled by user.")
            return

        # Build a small results table dialog.
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Algorithm comparison")
        dlg.setModal(True)
        dlg.resize(560, 320)
        layout = QtWidgets.QVBoxLayout(dlg)

        table = QtWidgets.QTableWidget(len(results), 6)
        table.setHorizontalHeaderLabels(
            ["Algorithm", "Clusters", "Silhouette", "Calinski-H", "Davies-B", "Runtime (s)"]
        )
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(True)
        table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)

        def _fmt(value):
            return "—" if value is None else f"{value:.3f}"

        for row, r in enumerate(results):
            cells = [
                r.name,
                str(r.n_clusters),
                _fmt(r.silhouette),
                _fmt(r.calinski_harabasz),
                _fmt(r.davies_bouldin),
                f"{r.runtime_seconds:.2f}",
            ]
            for col, text in enumerate(cells):
                item = QtWidgets.QTableWidgetItem(text)
                if col > 0:
                    item.setTextAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
                table.setItem(row, col, item)
        table.resizeColumnsToContents()
        layout.addWidget(table)

        best = _best(results)
        hint = QtWidgets.QLabel(
            f"Best by silhouette: <b>{best}</b>" if best
            else "No algorithm produced a valid silhouette score."
        )
        hint.setProperty("role", "muted")
        layout.addWidget(hint)

        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        if best:
            apply_btn = QtWidgets.QPushButton(f"Use {best}")
            apply_btn.setProperty("primary", "true")
            apply_btn.clicked.connect(lambda: (self._set_algorithm(best), dlg.accept()))
            buttons.addButton(apply_btn, QtWidgets.QDialogButtonBox.AcceptRole)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)
        dlg.exec()

        self.log_msg(f"Algorithm comparison: best by silhouette = {best or '(none)'}")

    # ------------------------------------------------------------------ #
    # Persistence: geometry / close                                      #
    # ------------------------------------------------------------------ #

    def _restore_geometry(self):
        geo = self.settings.get("geometry_main")
        if geo:
            try:
                self.restoreGeometry(QtCore.QByteArray.fromBase64(geo.encode("ascii")))
            except Exception:
                pass
        state = self.settings.get("window_state")
        if state:
            try:
                self.restoreState(QtCore.QByteArray.fromBase64(state.encode("ascii")))
            except Exception:
                pass
        # Guard: ensure the window is on a visible screen and not too small.
        self._clamp_to_visible_screen()

    def _clamp_to_visible_screen(self) -> None:
        """Ensure the window geometry fits within available screens.

        Handles cases where saved geometry references a monitor that is
        no longer connected, or the resolution has changed.
        """
        screen = self.screen()
        if screen is None:
            return
        avail = screen.availableGeometry()
        geo = self.geometry()
        # Enforce minimum usable size.
        w = max(geo.width(), self.minimumWidth() or 1040)
        h = max(geo.height(), self.minimumHeight() or 740)
        # Clamp position so the window is at least partially visible.
        x = max(avail.left(), min(geo.x(), avail.right() - 200))
        y = max(avail.top(), min(geo.y(), avail.bottom() - 100))
        # Don't exceed screen size.
        w = min(w, avail.width())
        h = min(h, avail.height())
        self.setGeometry(x, y, w, h)

    def closeEvent(self, event):
        try:
            self.settings["geometry_main"] = bytes(self.saveGeometry().toBase64()).decode("ascii")
            self.settings["window_state"] = bytes(self.saveState().toBase64()).decode("ascii")
            if self.wordcloud_builder is not None:
                self.settings["geometry_wordcloud"] = bytes(self.wordcloud_builder.saveGeometry().toBase64()).decode("ascii")
            self.settings["last_column"] = self.current_column_name()
            self.settings["last_algorithm"] = self.current_algorithm()
            app_settings.save(self.settings)
        except Exception:
            pass
        super().closeEvent(event)



def main():
    # HiDPI bootstrap — must run before QApplication is created.
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling, True)
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps, True)
    try:
        QtGui.QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
            QtCore.Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        # Older Qt builds — fall back silently.
        pass

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)

    # ── Application identity (used by QSettings, macOS About menu, etc.) ──
    app.setApplicationName("Text Analyzer Pro")
    app.setApplicationVersion("2.0")
    app.setOrganizationName("Aneek Hait")
    app.setOrganizationDomain("aneekhait.dev")

    # ── Platform tweaks ──
    if sys.platform == "darwin":
        # macOS: ensure the app name appears in the menu-bar About item.
        app.setApplicationDisplayName("Text Analyzer Pro")
        # Enable native event processing for Retina displays.
        app.setAttribute(QtCore.Qt.AA_DontCreateNativeWidgetSiblings, True)
    elif sys.platform == "linux":
        # Set desktop filename for proper taskbar grouping on Wayland/X11.
        try:
            app.setDesktopFileName("text-analyzer-pro")
        except AttributeError:
            pass

    icon = _load_app_icon()
    if not icon.isNull():
        app.setWindowIcon(icon)

    settings = app_settings.load()
    requested_theme = settings.get("theme", "system")
    if requested_theme == "system":
        mode = ThemeManager.detect_system_mode()
    else:
        mode = requested_theme
    theme_manager = ThemeManager(app, mode=mode)
    theme_manager.apply()

    window = ClusterGUI(theme_manager=theme_manager)
    window.showMaximized()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
