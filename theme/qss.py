"""QSS stylesheet builder. Pure function of token dict."""
from .tokens import get_tokens
from .assets import checkmark_url, chevron_down_url


def build_qss(mode: str = "light") -> str:
    t = get_tokens(mode)
    chevron_url = chevron_down_url(t["text_muted"])
    check_url = checkmark_url("#ffffff", 14)
    return f"""
/* ========== Base ========== */
* {{
    font-family: {t['font_family']};
    font-size: {t['font_size_md']};
    color: {t['text']};
}}
QWidget {{
    background-color: {t['bg']};
    color: {t['text']};
}}
QMainWindow, QDialog {{
    background-color: {t['bg']};
}}

/* ========== Sidebar ========== */
QFrame#Sidebar {{
    background-color: {t['sidebar_bg']};
    border: none;
}}
QFrame#Sidebar QLabel {{
    color: {t['sidebar_fg']};
    background: transparent;
}}
QLabel#SidebarBrand {{
    color: {t['sidebar_fg_active']};
    font-size: {t['font_size_lg']};
    font-weight: 600;
    padding: 4px 0px;
}}
QLabel#SidebarTagline {{
    color: {t['sidebar_fg']};
    font-size: {t['font_size_sm']};
    letter-spacing: 1px;
}}
QLabel#SidebarVersion {{
    color: {t['sidebar_fg']};
    font-size: {t['font_size_xs']};
    background-color: rgba(255, 255, 255, 0.08);
    border-radius: 8px;
    padding: 2px 10px;
}}
QToolButton#SidebarNav {{
    background-color: transparent;
    color: {t['sidebar_fg']};
    border: none;
    border-left: 3px solid transparent;
    border-radius: 6px;
    padding: 11px 14px;
    margin: 2px 4px;
    text-align: left;
    font-size: {t['font_size_md']};
    font-weight: 500;
}}
QToolButton#SidebarNav:hover {{
    background-color: {t['sidebar_active_bg']};
    color: {t['sidebar_fg_active']};
    border-left: 3px solid rgba(255, 255, 255, 0.15);
}}
QToolButton#SidebarNav:checked {{
    background-color: {t['sidebar_active_bg']};
    color: {t['sidebar_fg_active']};
    border-left: 3px solid {t['sidebar_accent']};
    font-weight: 600;
}}
QFrame#SidebarDivider {{
    background-color: {t['border_strong']};
    max-height: 1px;
    min-height: 1px;
    border: none;
    margin: 8px 12px;
}}

/* ========== Top bar ========== */
QFrame#TopBar {{
    background-color: {t['bg_elev']};
    border-bottom: 1px solid {t['divider']};
}}
QLabel#PageTitle {{
    font-size: {t['font_size_xl']};
    font-weight: 600;
    color: {t['text']};
}}
QLabel#PageSubtitle {{
    color: {t['text_muted']};
    font-size: {t['font_size_md']};
}}

/* ========== Action bar ========== */
QFrame#ActionBar {{
    background-color: {t['bg_elev']};
    border-top: 1px solid {t['divider']};
    min-height: 52px;
}}

/* ========== Data source panel ========== */
QFrame#DropZone {{
    background-color: {t['accent_subtle']};
    border: 2px dashed {t['border_strong']};
    border-radius: {t['radius_lg']};
}}
QFrame#DropZone:hover {{
    border-color: {t['accent']};
    background-color: {t['accent_soft']};
    border-style: solid;
}}
QFrame#DropZone[dragging="true"] {{
    border-color: {t['accent']};
    background-color: {t['accent_soft']};
    border-width: 3px;
    border-style: solid;
}}
QFrame#DropZone QLabel {{
    background: transparent;
}}
QLabel#DropZoneTitle {{
    font-size: {t['font_size_lg']};
    font-weight: 600;
    color: {t['text']};
}}
QLabel#DropZoneSubtitle {{
    color: {t['text_muted']};
    font-size: {t['font_size_sm']};
}}
QFrame#DataSourceSummary {{
    background-color: {t['surface_alt']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_sm']};
}}
QFrame#DataSourceSummary QLabel {{
    background: transparent;
}}
QLabel#DataSourceFile {{
    font-weight: 600;
    color: {t['text']};
}}
QTableView#PreviewTable {{
    background-color: {t['surface']};
    alternate-background-color: {t['surface_alt']};
    gridline-color: {t['divider']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_sm']};
}}
QTableView#PreviewTable QHeaderView::section {{
    background-color: {t['surface_alt']};
    color: {t['text_muted']};
    padding: 4px 8px;
    border: none;
    border-bottom: 1px solid {t['divider']};
    border-right: 1px solid {t['divider']};
    font-weight: 600;
}}

/* ========== Cards / GroupBoxes ========== */
QGroupBox {{
    background-color: {t['surface']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_md']};
    margin-top: 14px;
    padding: 12px 12px 10px 12px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 6px;
    left: 10px;
    color: {t['text']};
}}
QFrame#Card {{
    background-color: {t['surface']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_md']};
}}

/* ========== Buttons ========== */
QPushButton {{
    background-color: {t['surface']};
    color: {t['text']};
    border: 1px solid {t['border_strong']};
    border-radius: {t['radius_sm']};
    padding: 6px 14px;
    min-height: 22px;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: {t['surface_alt']};
    border-color: {t['accent']};
}}
QPushButton:pressed {{
    background-color: {t['accent_soft']};
}}
QPushButton:disabled {{
    color: {t['text_subtle']};
    background-color: {t['surface_alt']};
    border-color: {t['border']};
}}

QPushButton[primary="true"] {{
    background-color: {t['accent']};
    color: #ffffff;
    border: 1px solid {t['accent']};
}}
QPushButton[primary="true"]:hover {{
    background-color: {t['accent_hover']};
    border-color: {t['accent_hover']};
}}
QPushButton[primary="true"]:pressed {{
    background-color: {t['accent_pressed']};
    border-color: {t['accent_pressed']};
}}
QPushButton[primary="true"]:disabled {{
    background-color: {t['border']};
    color: {t['text_subtle']};
    border-color: {t['border']};
}}

QPushButton[danger="true"] {{
    background-color: {t['danger']};
    color: #ffffff;
    border: 1px solid {t['danger']};
}}
QPushButton[danger="true"]:hover {{
    background-color: {t['danger_hover']};
    border-color: {t['danger_hover']};
}}

QPushButton[flat="true"] {{
    background-color: transparent;
    border: 1px solid transparent;
    color: {t['accent']};
    padding: 5px 10px;
}}
QPushButton[flat="true"]:hover {{
    background-color: {t['accent_soft']};
    border-color: {t['accent']};
    border-radius: {t['radius_sm']};
}}
QPushButton[flat="true"]:pressed {{
    background-color: {t['accent_soft']};
    color: {t['accent_pressed']};
}}

/* ========== Inputs ========== */
QLineEdit, QPlainTextEdit, QTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
    background-color: {t['input_bg']};
    color: {t['text']};
    border: 1px solid {t['input_border']};
    border-radius: {t['radius_sm']};
    padding: 5px 8px;
    selection-background-color: {t['selection_bg']};
    selection-color: {t['selection_fg']};
}}
QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus,
QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {{
    border: 1.5px solid {t['input_focus']};
}}
QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QComboBox:disabled {{
    color: {t['text_subtle']};
    background-color: {t['surface_alt']};
}}
QPlainTextEdit, QTextEdit {{
    font-family: {t['font_mono']};
    font-size: {t['font_size_md']};
}}
QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: center right;
    border: none;
    width: 22px;
}}
QComboBox::down-arrow {{
    image: url("{chevron_url}");
    width: 10px;
    height: 10px;
    margin-right: 8px;
}}
QComboBox QAbstractItemView {{
    background-color: {t['bg_elev']};
    color: {t['text']};
    border: 1px solid {t['border_strong']};
    selection-background-color: {t['accent_soft']};
    selection-color: {t['text']};
    outline: none;
}}

/* ========== Checkbox / Radio ========== */
QCheckBox, QRadioButton {{
    spacing: 8px;
    background: transparent;
}}
QCheckBox::indicator, QRadioButton::indicator {{
    width: 16px;
    height: 16px;
    border: 1px solid {t['border_strong']};
    background: {t['input_bg']};
    border-radius: 3px;
}}
QRadioButton::indicator {{
    border-radius: 8px;
}}
QCheckBox::indicator:hover, QRadioButton::indicator:hover {{
    border: 1px solid {t['accent']};
}}
QCheckBox::indicator:checked {{
    background: {t['accent']};
    border: 1px solid {t['accent']};
    image: url("{check_url}");
}}
QRadioButton::indicator:checked {{
    background: {t['accent']};
    border: 1px solid {t['accent']};
}}

/* ========== Tabs (kept for sub-tabs if present) ========== */
QTabWidget::pane {{
    border: 1px solid {t['border']};
    border-radius: {t['radius_md']};
    background-color: {t['surface']};
    top: -1px;
}}
QTabBar::tab {{
    background-color: transparent;
    color: {t['text_muted']};
    padding: 8px 16px;
    border: 1px solid transparent;
    border-bottom: none;
    border-top-left-radius: {t['radius_sm']};
    border-top-right-radius: {t['radius_sm']};
    margin-right: 2px;
}}
QTabBar::tab:hover {{
    color: {t['text']};
}}
QTabBar::tab:selected {{
    background-color: {t['surface']};
    color: {t['text']};
    border: 1px solid {t['border']};
    border-bottom: 1px solid {t['surface']};
    font-weight: 600;
}}

/* ========== Tables ========== */
QTableWidget, QTableView {{
    background-color: {t['surface']};
    alternate-background-color: {t['surface_alt']};
    gridline-color: {t['divider']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_md']};
    selection-background-color: {t['accent_soft']};
    selection-color: {t['text']};
}}
QHeaderView::section {{
    background-color: {t['surface_alt']};
    color: {t['text_muted']};
    padding: 6px 8px;
    border: none;
    border-right: 1px solid {t['divider']};
    border-bottom: 1px solid {t['border']};
    font-weight: 600;
}}
QTableCornerButton::section {{
    background-color: {t['surface_alt']};
    border: none;
    border-bottom: 1px solid {t['border']};
}}

/* ========== Lists ========== */
QListWidget, QListView, QTreeView {{
    background-color: {t['surface']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_md']};
    padding: 2px;
    outline: none;
}}
QListWidget::item, QListView::item {{
    padding: 5px 6px;
    border-radius: 3px;
}}
QListWidget::item:selected, QListView::item:selected {{
    background-color: {t['accent_soft']};
    color: {t['text']};
}}
QListWidget::item:hover, QListView::item:hover {{
    background-color: {t['surface_alt']};
}}

/* ========== Progress bar ========== */
QProgressBar {{
    background-color: {t['surface_alt']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_sm']};
    text-align: center;
    color: {t['text']};
    height: 14px;
}}
QProgressBar::chunk {{
    background-color: {t['accent']};
    border-radius: 3px;
}}

/* ========== Status bar ========== */
QStatusBar {{
    background-color: {t['bg_elev']};
    color: {t['text_muted']};
    border-top: 1px solid {t['divider']};
}}
QStatusBar::item {{
    border: none;
}}
QStatusBar QLabel {{
    color: {t['text_muted']};
    padding: 0 6px;
}}

/* ========== Menu ========== */
QMenuBar {{
    background-color: {t['bg_elev']};
    border-bottom: 1px solid {t['divider']};
    color: {t['text']};
}}
QMenuBar::item {{
    background: transparent;
    padding: 6px 12px;
}}
QMenuBar::item:selected {{
    background-color: {t['accent_soft']};
}}
QMenu {{
    background-color: {t['bg_elev']};
    border: 1px solid {t['border_strong']};
    border-radius: {t['radius_sm']};
    padding: 4px;
}}
QMenu::item {{
    padding: 6px 24px 6px 16px;
    border-radius: 3px;
}}
QMenu::item:selected {{
    background-color: {t['accent_soft']};
}}
QMenu::separator {{
    height: 1px;
    background: {t['divider']};
    margin: 4px 6px;
}}

/* ========== Tooltip ========== */
QToolTip {{
    background-color: {t['tooltip_bg']};
    color: {t['tooltip_fg']};
    border: 1px solid {t['border_strong']};
    border-radius: {t['radius_sm']};
    padding: 5px 8px;
}}

/* ========== Scrollbars ========== */
QScrollBar:vertical {{
    background: {t['scroll_bg']};
    width: 12px;
    margin: 0;
    border: none;
}}
QScrollBar::handle:vertical {{
    background: {t['scroll_thumb']};
    min-height: 30px;
    border-radius: 6px;
    margin: 2px;
}}
QScrollBar::handle:vertical:hover {{
    background: {t['scroll_thumb_hover']};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
    background: none;
    border: none;
}}
QScrollBar:horizontal {{
    background: {t['scroll_bg']};
    height: 12px;
    margin: 0;
    border: none;
}}
QScrollBar::handle:horizontal {{
    background: {t['scroll_thumb']};
    min-width: 30px;
    border-radius: 6px;
    margin: 2px;
}}
QScrollBar::handle:horizontal:hover {{
    background: {t['scroll_thumb_hover']};
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0;
    background: none;
    border: none;
}}

/* ========== Splitter ========== */
QSplitter::handle {{
    background-color: {t['divider']};
}}
QSplitter::handle:horizontal {{
    width: 3px;
}}
QSplitter::handle:vertical {{
    height: 3px;
}}

/* ========== ScrollArea ========== */
QScrollArea {{
    background-color: transparent;
    border: none;
}}

/* ========== Custom: cluster card ========== */
QFrame#ClusterCard {{
    background-color: {t['surface']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_md']};
    padding: 6px;
}}
QLabel#ClusterChip {{
    background-color: {t['accent_soft']};
    color: {t['accent']};
    border-radius: 9px;
    padding: 2px 10px;
    font-weight: 600;
    font-size: {t['font_size_sm']};
}}
QLabel#KeywordPill {{
    background-color: {t['surface_alt']};
    color: {t['text_muted']};
    border-radius: 8px;
    padding: 2px 8px;
    font-size: {t['font_size_sm']};
}}

/* ========== Custom: toast ========== */
QFrame#Toast {{
    background-color: {t['bg_elev']};
    color: {t['text']};
    border: 1px solid {t['border_strong']};
    border-radius: {t['radius_md']};
    padding: 10px 14px;
}}
QFrame#Toast[level="success"] {{
    border-left: 3px solid {t['success']};
}}
QFrame#Toast[level="warning"] {{
    border-left: 3px solid {t['warning']};
}}
QFrame#Toast[level="danger"] {{
    border-left: 3px solid {t['danger']};
}}
QFrame#Toast[level="info"] {{
    border-left: 3px solid {t['accent']};
}}

/* ========== Section headers ========== */
QLabel[role="h1"] {{
    font-size: {t['font_size_xxl']};
    font-weight: 600;
    color: {t['text']};
}}
QLabel[role="h2"] {{
    font-size: {t['font_size_xl']};
    font-weight: 600;
    color: {t['text']};
}}
QLabel[role="muted"] {{
    color: {t['text_muted']};
}}
QLabel[role="subtle"] {{
    color: {t['text_subtle']};
    font-size: {t['font_size_sm']};
}}

/* ========== Workspace tabs (top-level QTabWidget) ========== */
QTabWidget#WorkspaceTabs::pane {{
    border: 1px solid {t['border']};
    border-top: 1px solid {t['border']};
    background-color: {t['bg']};
    top: -1px;
}}
QTabWidget#WorkspaceTabs QTabBar {{
    qproperty-drawBase: 0;
    background-color: {t['bg']};
}}
QTabWidget#WorkspaceTabs QTabBar::tab {{
    background-color: {t['surface_alt']};
    color: {t['text_muted']};
    padding: 6px 14px;
    margin-right: 2px;
    border: 1px solid {t['border']};
    border-bottom: none;
    border-top-left-radius: {t['radius_md']};
    border-top-right-radius: {t['radius_md']};
    min-width: 120px;
    font-size: {t['font_size_md']};
}}
QTabWidget#WorkspaceTabs QTabBar::tab:hover {{
    background-color: {t['surface']};
    color: {t['text']};
}}
QTabWidget#WorkspaceTabs QTabBar::tab:selected {{
    background-color: {t['bg']};
    color: {t['text']};
    border-bottom: 1px solid {t['bg']};
}}
QTabWidget#WorkspaceTabs QTabBar::close-button {{
    subcontrol-position: right;
    margin: 4px;
}}

/* ========== Inner tabs (per-tab Setup/Cleaning/Results) ========== */
QTabWidget#InnerTabs::pane {{
    border: none;
    background-color: transparent;
    top: -1px;
}}
QTabWidget#InnerTabs QTabBar::tab {{
    background-color: transparent;
    color: {t['text_muted']};
    padding: 8px 16px;
    border: none;
    border-bottom: 2px solid transparent;
    margin-right: 4px;
    font-size: {t['font_size_md']};
    font-weight: 500;
}}
QTabWidget#InnerTabs QTabBar::tab:hover {{
    color: {t['text']};
}}
QTabWidget#InnerTabs QTabBar::tab:selected {{
    color: {t['accent']};
    border-bottom: 2px solid {t['accent']};
}}

/* ========== QDockWidget ========== */
QDockWidget {{
    color: {t['text']};
    titlebar-close-icon: none;
    titlebar-normal-icon: none;
}}
QDockWidget::title {{
    background-color: {t['surface_alt']};
    color: {t['text_muted']};
    padding: 6px 10px;
    border: 1px solid {t['border']};
    border-bottom: none;
    text-align: left;
    font-size: {t['font_size_sm']};
    font-weight: 600;
}}
QDockWidget > QWidget {{
    background-color: {t['bg_elev']};
    border: 1px solid {t['border']};
}}

/* ========== QSplitter ========== */
QSplitter::handle {{
    background-color: {t['border']};
}}
QSplitter::handle:horizontal {{
    width: 1px;
}}
QSplitter::handle:vertical {{
    height: 1px;
}}
QSplitter::handle:hover {{
    background-color: {t['accent']};
}}

/* ========== Cards & sections ========== */
QFrame#Card, QFrame[role="card"] {{
    background-color: {t['bg_elev']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_lg']};
}}
QFrame#SectionCard {{
    background-color: {t['bg_elev']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_lg']};
}}

/* ========== Empty state ========== */
QWidget#EmptyState {{
    background-color: {t['bg']};
}}
QLabel#EmptyStateTitle {{
    color: {t['text']};
    font-size: {t['font_size_xl']};
    font-weight: 600;
}}
QLabel#EmptyStateBody {{
    color: {t['text_muted']};
    font-size: {t['font_size_md']};
}}

/* ========== Error card ========== */
QFrame#ErrorCard {{
    background-color: {t['danger_soft']};
    border: 1px solid {t['danger']};
    border-radius: {t['radius_md']};
    padding: 12px;
}}
QFrame#ErrorCard QLabel#ErrorCardTitle {{
    color: {t['danger']};
    font-weight: 600;
    font-size: {t['font_size_md']};
}}
QFrame#ErrorCard QLabel#ErrorCardBody {{
    color: {t['text']};
    font-size: {t['font_size_sm']};
}}

/* ========== Metric card ========== */
QFrame#MetricCard {{
    background-color: {t['bg_elev']};
    border: 1px solid {t['border']};
    border-radius: {t['radius_lg']};
    padding: 14px 16px;
}}
QFrame#MetricCard QLabel[role="metric_label"] {{
    color: {t['text_muted']};
    font-size: {t['font_size_sm']};
    font-weight: 500;
}}
QFrame#MetricCard QLabel[role="metric_value"] {{
    color: {t['text']};
    font-size: {t['font_size_xxl']};
    font-weight: 600;
}}

/* ========== KBD chip ========== */
QLabel[role="kbd"] {{
    background-color: {t['kbd_bg']};
    color: {t['kbd_fg']};
    border: 1px solid {t['kbd_border']};
    border-radius: {t['radius_sm']};
    padding: 1px 6px;
    font-family: {t['font_mono']};
    font-size: {t['font_size_xs']};
}}

/* ========== Focus ring (overrides default) ========== */
QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus,
QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QAbstractSpinBox:focus {{
    border: 1px solid {t['focus_ring']};
}}
QPushButton:focus {{
    outline: none;
    border: 1px solid {t['focus_ring']};
}}

/* ========== App header / toolbar ========== */
QToolBar#AppHeader {{
    background-color: {t['bg_elev']};
    border-bottom: 1px solid {t['border']};
    spacing: 6px;
    padding: 6px 10px;
}}
QToolBar#AppHeader QToolButton {{
    background-color: transparent;
    color: {t['text']};
    border: 1px solid transparent;
    border-radius: {t['radius_md']};
    padding: 5px 10px;
}}
QToolBar#AppHeader QToolButton:hover {{
    background-color: {t['surface_alt']};
    border-color: {t['border']};
}}
QToolBar#AppHeader QToolButton:pressed {{
    background-color: {t['accent_soft']};
}}
QToolBar#AppHeader QToolButton:checked {{
    background-color: {t['accent_soft']};
    color: {t['accent']};
    border-color: {t['accent']};
}}
"""
