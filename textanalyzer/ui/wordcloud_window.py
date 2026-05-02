"""Word Cloud dialog — PySide6 port of the distributable's WordCloudDialog.

Replicates the exact layout: toolbar -> 3-column body (sidebar | preview | word
counts) -> status bar, with the same settings groups and features.
"""

from __future__ import annotations

import io
import os
import re
from collections import Counter

from PIL import Image
from PIL.ImageQt import ImageQt
from PySide6 import QtCore, QtGui, QtWidgets

from cluster_tool import coerce_text_column
from wordcloud_tool import (
    WORDCLOUD_BACKGROUNDS,
    WORDCLOUD_COLOR_SCHEMES,
    WORDCLOUD_SHAPES,
    create_shape_mask,
    discover_system_fonts,
    generate_wordcloud,
    load_custom_mask,
    save_wordcloud,
    wordcloud_to_image,
)

from ..utils.icons import _icon


_LIVE_RENDER_DEBOUNCE_MS = 350


class WordCloudDialog(QtWidgets.QDialog):
    """Word Cloud generation dialog matching the distributable layout."""

    def __init__(self, parent, texts: list, column_name: str, dataframe=None):
        super().__init__(parent)
        self.texts = texts
        self.column_name = column_name
        self.dataframe = dataframe
        self.wc = None
        self.word_frequencies: dict = {}
        self.preview_scale = 1.0
        self.custom_stopwords: set = set()
        self.custom_mask_path: str | None = None
        self.actual_word_counts: Counter = Counter()
        self.total_word_count = 0

        self.setWindowTitle(f"Word Cloud \u2014 {column_name}")
        self.resize(1200, 750)
        self.setMinimumSize(1000, 600)
        # QDialog hides minimize/maximize by default. Re-enable them so
        # Wordcloud Studio behaves like a regular top-level window.
        self.setWindowFlags(
            self.windowFlags()
            | QtCore.Qt.WindowMinimizeButtonHint
            | QtCore.Qt.WindowMaximizeButtonHint
        )

        self._font_paths = discover_system_fonts()

        # Debounce timer for live rendering. Settings emit many rapid changes
        # (slider drags, spinbox tap-and-holds); coalesce them into one render.
        self._debounce_timer = QtCore.QTimer(self)
        self._debounce_timer.setSingleShot(True)
        self._debounce_timer.setInterval(_LIVE_RENDER_DEBOUNCE_MS)
        self._debounce_timer.timeout.connect(self._generate_preview)

        self._build_ui()
        self._generate_preview()

    def _schedule_regen(self) -> None:
        """Restart the debounce timer; the actual render fires when it expires."""
        self._debounce_timer.start()

    # ================================================================
    # UI construction
    # ================================================================

    def _build_ui(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- TOP TOOLBAR ------------------------------------------------
        toolbar = QtWidgets.QFrame()
        toolbar.setObjectName("Toolbar")
        tb = QtWidgets.QHBoxLayout(toolbar)
        tb.setContentsMargins(10, 6, 10, 6)
        tb.setSpacing(8)

        self.render_status = QtWidgets.QLabel("Live preview")
        self.render_status.setProperty("role", "muted")
        self.render_status.setStyleSheet("font-size: 9pt; padding-left: 4px;")
        tb.addWidget(self.render_status)

        sep1 = QtWidgets.QFrame()
        sep1.setFrameShape(QtWidgets.QFrame.VLine)
        sep1.setFrameShadow(QtWidgets.QFrame.Sunken)
        tb.addWidget(sep1)

        copy_btn = QtWidgets.QPushButton("Copy to Clipboard")
        copy_btn.setIcon(_icon("fa5s.clipboard"))
        copy_btn.clicked.connect(self._copy_to_clipboard)
        tb.addWidget(copy_btn)

        save_btn = QtWidgets.QPushButton("Save Image\u2026")
        save_btn.setIcon(_icon("fa5s.save"))
        save_btn.clicked.connect(self._save_image)
        tb.addWidget(save_btn)

        tb.addStretch(1)
        root.addWidget(toolbar)

        sep_top = QtWidgets.QFrame()
        sep_top.setFrameShape(QtWidgets.QFrame.HLine)
        sep_top.setFrameShadow(QtWidgets.QFrame.Sunken)
        root.addWidget(sep_top)

        # ---- BODY --- three columns via QSplitter -----------------------
        body = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        body.setChildrenCollapsible(False)
        body.setHandleWidth(6)
        root.addWidget(body, 1)

        # ---- LEFT SIDEBAR -----------------------------------------------
        sidebar_scroll = QtWidgets.QScrollArea()
        sidebar_scroll.setWidgetResizable(True)
        sidebar_scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        sidebar_scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        sidebar_scroll.setMinimumWidth(260)
        sidebar_scroll.setMaximumWidth(340)

        sidebar = QtWidgets.QWidget()
        sb = QtWidgets.QVBoxLayout(sidebar)
        sb.setContentsMargins(14, 12, 10, 14)
        sb.setSpacing(12)

        def _make_form(group: QtWidgets.QGroupBox) -> QtWidgets.QFormLayout:
            form = QtWidgets.QFormLayout(group)
            # Tight L/R margins because QGroupBox already adds 12 px of inner
            # padding via the global QSS — compounding both clips fields.
            form.setContentsMargins(2, 8, 2, 8)
            form.setHorizontalSpacing(10)
            form.setVerticalSpacing(10)
            form.setLabelAlignment(QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
            form.setFieldGrowthPolicy(QtWidgets.QFormLayout.AllNonFixedFieldsGrow)
            # When the sidebar is narrow, drop the label onto its own line
            # above the field instead of squeezing both side-by-side. Prevents
            # labels from being clipped when the user drags the splitter
            # narrower.
            form.setRowWrapPolicy(QtWidgets.QFormLayout.WrapLongRows)
            return form

        def _shrinkable_combo() -> QtWidgets.QComboBox:
            """QComboBox whose minimumSizeHint is short enough that long items
            don't push the sidebar width past its viewport. Items still render
            in full inside the popup."""
            combo = QtWidgets.QComboBox()
            combo.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToMinimumContentsLengthWithIcon)
            combo.setMinimumContentsLength(8)
            combo.setSizePolicy(
                QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed
            )
            return combo

        # ---- Source group (column selector) -----------------------------
        grp_source = QtWidgets.QGroupBox("\U0001f4c1  Source")
        gsr = _make_form(grp_source)

        self.column_combo = _shrinkable_combo()
        self._populate_column_combo()
        self.column_combo.currentTextChanged.connect(self._on_column_change)
        gsr.addRow("Column", self.column_combo)

        sb.addWidget(grp_source)
        # Hide entirely when no dataframe was supplied (legacy callers).
        grp_source.setVisible(self.dataframe is not None and self.column_combo.count() > 0)

        # ---- Appearance group -------------------------------------------
        grp_appear = QtWidgets.QGroupBox("\U0001f3a8  Appearance")
        ga = _make_form(grp_appear)

        self.font_combo = _shrinkable_combo()
        self.font_combo.addItems(list(self._font_paths.keys()))
        ga.addRow("Font", self.font_combo)

        self.color_combo = _shrinkable_combo()
        self.color_combo.addItems(list(WORDCLOUD_COLOR_SCHEMES.keys()))
        self.color_combo.setCurrentText("Corporate Blue")
        ga.addRow("Colors", self.color_combo)

        self.bg_combo = _shrinkable_combo()
        self.bg_combo.addItems(list(WORDCLOUD_BACKGROUNDS.keys()))
        ga.addRow("Background", self.bg_combo)

        sb.addWidget(grp_appear)

        # ---- Shape group ------------------------------------------------
        grp_shape = QtWidgets.QGroupBox("\u2726  Shape")
        gs = _make_form(grp_shape)

        self.shape_combo = _shrinkable_combo()
        shape_values = list(WORDCLOUD_SHAPES.keys()) + ["Custom Image\u2026"]
        self.shape_combo.addItems(shape_values)
        self.shape_combo.currentTextChanged.connect(self._on_shape_change)
        gs.addRow("Shape", self.shape_combo)

        self.custom_shape_label = QtWidgets.QLabel("")
        self.custom_shape_label.setProperty("role", "muted")
        self.custom_shape_label.setWordWrap(True)
        self.custom_shape_label.setStyleSheet("font-size: 8pt;")
        # Hidden until a custom image is picked \u2014 saves a wasted empty row.
        self.custom_shape_label.hide()
        gs.addRow(self.custom_shape_label)

        sb.addWidget(grp_shape)

        # ---- Typography group -------------------------------------------
        grp_typo = QtWidgets.QGroupBox("Aa  Typography")
        gt = _make_form(grp_typo)

        self.max_words_spin = QtWidgets.QSpinBox()
        self.max_words_spin.setRange(50, 500)
        self.max_words_spin.setValue(200)
        self.max_words_spin.setSingleStep(10)
        gt.addRow("Max words", self.max_words_spin)

        self.min_font_spin = QtWidgets.QSpinBox()
        self.min_font_spin.setRange(4, 40)
        self.min_font_spin.setValue(10)
        self.min_font_spin.setSingleStep(2)
        gt.addRow("Min font", self.min_font_spin)

        self.max_font_spin = QtWidgets.QSpinBox()
        self.max_font_spin.setRange(0, 500)
        self.max_font_spin.setValue(0)
        self.max_font_spin.setSingleStep(10)
        self.max_font_spin.setToolTip("0 = auto")
        gt.addRow("Max font", self.max_font_spin)

        sb.addWidget(grp_typo)

        # ---- Scaling group (Size proportional to Count) -----------------
        grp_scale = QtWidgets.QGroupBox("\u2696  Size \u221d Count")
        scale_layout = QtWidgets.QVBoxLayout(grp_scale)
        scale_layout.setContentsMargins(2, 8, 2, 8)
        scale_layout.setSpacing(8)

        self.rel_scale_label = QtWidgets.QLabel("0.0  (rank only)")
        self.rel_scale_label.setStyleSheet("font-size: 9pt;")
        scale_layout.addWidget(self.rel_scale_label)

        self.rel_scale_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.rel_scale_slider.setRange(0, 100)
        self.rel_scale_slider.setValue(0)
        self.rel_scale_slider.valueChanged.connect(self._on_rel_scale_change)
        scale_layout.addWidget(self.rel_scale_slider)

        hint_row = QtWidgets.QHBoxLayout()
        lbl_rank = QtWidgets.QLabel("rank")
        lbl_rank.setProperty("role", "muted")
        lbl_rank.setStyleSheet("font-size: 8pt;")
        lbl_count = QtWidgets.QLabel("count")
        lbl_count.setProperty("role", "muted")
        lbl_count.setStyleSheet("font-size: 8pt;")
        lbl_count.setAlignment(QtCore.Qt.AlignRight)
        hint_row.addWidget(lbl_rank)
        hint_row.addStretch(1)
        hint_row.addWidget(lbl_count)
        scale_layout.addLayout(hint_row)

        sb.addWidget(grp_scale)

        # ---- Stopwords group --------------------------------------------
        grp_sw = QtWidgets.QGroupBox("\U0001f6ab  Stopwords")
        sw_layout = QtWidgets.QVBoxLayout(grp_sw)
        sw_layout.setContentsMargins(2, 8, 2, 8)
        sw_layout.setSpacing(8)

        self.stopwords_count_label = QtWidgets.QLabel("0 custom words excluded")
        self.stopwords_count_label.setStyleSheet("font-size: 9pt;")
        sw_layout.addWidget(self.stopwords_count_label)

        self.stopwords_preview = QtWidgets.QLabel("")
        self.stopwords_preview.setProperty("role", "muted")
        self.stopwords_preview.setWordWrap(True)
        self.stopwords_preview.setStyleSheet("font-size: 8pt;")
        sw_layout.addWidget(self.stopwords_preview)

        sw_btn_row = QtWidgets.QHBoxLayout()
        sw_btn_row.setSpacing(8)
        edit_sw_btn = QtWidgets.QPushButton("Edit")
        edit_sw_btn.clicked.connect(self._edit_stopwords)
        clear_sw_btn = QtWidgets.QPushButton("Clear")
        clear_sw_btn.clicked.connect(self._clear_stopwords)
        sw_btn_row.addWidget(edit_sw_btn, 1)
        sw_btn_row.addWidget(clear_sw_btn, 1)
        sw_layout.addLayout(sw_btn_row)

        sb.addWidget(grp_sw)

        sb.addStretch(1)
        sidebar_scroll.setWidget(sidebar)
        body.addWidget(sidebar_scroll)

        # ---- CENTER --- Preview canvas ----------------------------------
        center = QtWidgets.QGroupBox("Preview")
        cl = QtWidgets.QVBoxLayout(center)
        cl.setContentsMargins(4, 4, 4, 4)

        self.preview_label = QtWidgets.QLabel("Generating\u2026")
        self.preview_label.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_label.setMinimumHeight(360)
        self.preview_label.setProperty("role", "muted")
        self.preview_label.setStyleSheet(
            "background: white; border: 1px solid rgba(120,120,120,0.18); border-radius: 6px;"
        )
        cl.addWidget(self.preview_label, 1)
        body.addWidget(center)

        # ---- RIGHT --- Word Counts panel --------------------------------
        right = QtWidgets.QGroupBox("Word Counts")
        rl = QtWidgets.QVBoxLayout(right)
        rl.setContentsMargins(6, 6, 6, 6)
        rl.setSpacing(4)
        right.setMinimumWidth(220)
        right.setMaximumWidth(280)

        # Search / filter
        filter_row = QtWidgets.QHBoxLayout()
        filter_row.setSpacing(4)
        f_icon = QtWidgets.QLabel("\U0001f50d")
        self.word_filter_edit = QtWidgets.QLineEdit()
        self.word_filter_edit.setPlaceholderText("Filter words\u2026")
        self.word_filter_edit.textChanged.connect(self._on_word_filter_change)
        filter_row.addWidget(f_icon)
        filter_row.addWidget(self.word_filter_edit, 1)
        rl.addLayout(filter_row)

        # Tree widget (word, count, %)
        self.word_tree = QtWidgets.QTreeWidget()
        self.word_tree.setHeaderLabels(["Word", "Count", "%"])
        self.word_tree.setRootIsDecorated(False)
        self.word_tree.setAlternatingRowColors(True)
        self.word_tree.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        header = self.word_tree.header()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)

        # Right-click context menu
        self.word_tree.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.word_tree.customContextMenuRequested.connect(self._show_word_context_menu)
        rl.addWidget(self.word_tree, 1)

        hint = QtWidgets.QLabel("Right-click a word \u2192 add to stopwords")
        hint.setProperty("role", "muted")
        hint.setStyleSheet("font-size: 8pt;")
        rl.addWidget(hint)

        body.addWidget(right)

        body.setStretchFactor(0, 0)
        body.setStretchFactor(1, 1)
        body.setStretchFactor(2, 0)
        body.setSizes([280, 700, 240])

        # ---- BOTTOM STATUS BAR ------------------------------------------
        sep_bot = QtWidgets.QFrame()
        sep_bot.setFrameShape(QtWidgets.QFrame.HLine)
        sep_bot.setFrameShadow(QtWidgets.QFrame.Sunken)
        root.addWidget(sep_bot)

        status_bar = QtWidgets.QFrame()
        status_bar.setObjectName("StatusBar")
        sbl = QtWidgets.QHBoxLayout(status_bar)
        sbl.setContentsMargins(10, 4, 10, 4)
        sbl.setSpacing(4)

        self.stats_label = QtWidgets.QLabel("")
        self.stats_label.setProperty("role", "muted")
        self.stats_label.setStyleSheet("font-size: 9pt;")
        sbl.addWidget(self.stats_label)

        root.addWidget(status_bar)

        # ---- Local QSS -------------------------------------------------
        self.setStyleSheet(
            self.styleSheet()
            + """
            QFrame#Toolbar { background: rgba(120,120,120,0.05); }
            QFrame#StatusBar { background: rgba(120,120,120,0.06); }
            QGroupBox { font-weight: 600; border: 1px solid rgba(120,120,120,0.22);
                        border-radius: 6px; margin-top: 10px; padding-top: 4px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
        """
        )

        # Wire all settings to the live-render debounce. Done after _build_ui
        # so every widget exists, and after the constructor's initial render
        # so the first generate isn't double-fired.
        self._connect_live_render_signals()

    def _connect_live_render_signals(self) -> None:
        for combo in (self.font_combo, self.color_combo, self.bg_combo):
            combo.currentTextChanged.connect(self._schedule_regen)
        for spin in (self.max_words_spin, self.min_font_spin, self.max_font_spin):
            spin.valueChanged.connect(self._schedule_regen)
        self.rel_scale_slider.valueChanged.connect(self._schedule_regen)
        # shape_combo's existing _on_shape_change handler covers the file
        # picker; piggyback a regen on top via a second connection. The branch
        # in _on_shape_change that resets to "Rectangle" already fires
        # currentTextChanged again, which restarts the debounce \u2014 that's fine.
        self.shape_combo.currentTextChanged.connect(self._schedule_regen)

    # ================================================================
    # Source column selector
    # ================================================================

    def _populate_column_combo(self) -> None:
        if self.dataframe is None:
            return
        # Prefer string/object columns, but fall back to all columns so users
        # can still pick a numeric column if they want.
        import pandas as pd

        cols = list(self.dataframe.columns)
        text_cols = [c for c in cols if pd.api.types.is_string_dtype(self.dataframe[c])]
        ordered = text_cols + [c for c in cols if c not in text_cols]
        self.column_combo.blockSignals(True)
        self.column_combo.clear()
        self.column_combo.addItems([str(c) for c in ordered])
        if self.column_name in [str(c) for c in ordered]:
            self.column_combo.setCurrentText(str(self.column_name))
        self.column_combo.blockSignals(False)

    def _on_column_change(self, name: str) -> None:
        if not name or self.dataframe is None or name not in self.dataframe.columns:
            return
        self.column_name = name
        self.texts = coerce_text_column(self.dataframe[name]).tolist()
        self.setWindowTitle(f"Word Cloud \u2014 {name}")
        self._schedule_regen()

    # ================================================================
    # Shape handling
    # ================================================================

    def _on_shape_change(self, text: str):
        if text == "Custom Image\u2026":
            path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self,
                "Select Shape Image",
                "",
                "Image files (*.png *.jpg *.jpeg *.bmp *.gif);;All files (*)",
            )
            if path:
                self.custom_mask_path = path
                self.custom_shape_label.setText(os.path.basename(path))
                self.custom_shape_label.show()
            else:
                self.shape_combo.blockSignals(True)
                self.shape_combo.setCurrentText("Rectangle")
                self.shape_combo.blockSignals(False)
                self.custom_shape_label.setText("")
                self.custom_shape_label.hide()
        else:
            self.custom_shape_label.setText("")
            self.custom_shape_label.hide()

    # ================================================================
    # Relative scaling slider
    # ================================================================

    def _on_rel_scale_change(self, value: int):
        v = value / 100.0
        if v >= 0.95:
            desc = "fully proportional"
        elif v >= 0.6:
            desc = "mostly proportional"
        elif v >= 0.3:
            desc = "balanced"
        elif v > 0.05:
            desc = "mostly rank-based"
        else:
            desc = "rank only"
        self.rel_scale_label.setText(f"{v:.1f}  ({desc})")

    # ================================================================
    # Stopwords
    # ================================================================

    def _update_stopwords_display(self):
        count = len(self.custom_stopwords)
        self.stopwords_count_label.setText(
            f"{count} custom word{'s' if count != 1 else ''} excluded"
        )
        if self.custom_stopwords:
            preview = ", ".join(sorted(self.custom_stopwords)[:5])
            if count > 5:
                preview += f", +{count - 5} more"
            self.stopwords_preview.setText(preview)
        else:
            self.stopwords_preview.setText("")

    def _clear_stopwords(self):
        if not self.custom_stopwords:
            return
        if (
            QtWidgets.QMessageBox.question(
                self,
                "Clear Stopwords",
                f"Remove all {len(self.custom_stopwords)} custom stopwords?",
            )
            == QtWidgets.QMessageBox.Yes
        ):
            self.custom_stopwords.clear()
            self._update_stopwords_display()
            self._generate_preview()

    def _edit_stopwords(self):
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Edit Stopwords")
        dlg.resize(350, 450)
        dlg.setModal(True)

        layout = QtWidgets.QVBoxLayout(dlg)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(8)

        layout.addWidget(QtWidgets.QLabel("Enter words to exclude (one per line):"))

        text_edit = QtWidgets.QPlainTextEdit()
        if self.custom_stopwords:
            text_edit.setPlainText("\n".join(sorted(self.custom_stopwords)))
        layout.addWidget(text_edit, 1)

        note = QtWidgets.QLabel("Note: Default English stopwords are always excluded.")
        note.setProperty("role", "muted")
        note.setStyleSheet("font-size: 8pt;")
        layout.addWidget(note)

        btn_row = QtWidgets.QHBoxLayout()
        clear_text_btn = QtWidgets.QPushButton("Clear Text")
        clear_text_btn.clicked.connect(text_edit.clear)
        save_btn = QtWidgets.QPushButton("Save")
        btn_row.addWidget(clear_text_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(save_btn)
        layout.addLayout(btn_row)

        def _save():
            text = text_edit.toPlainText()
            words = re.findall(r"\b[a-zA-Z0-9]+\b", text.lower())
            self.custom_stopwords = set(words)
            self._update_stopwords_display()
            dlg.accept()
            self._generate_preview()

        save_btn.clicked.connect(_save)
        dlg.exec()

    # ================================================================
    # Word counts (right panel)
    # ================================================================

    def _update_word_counts(self):
        self.word_tree.clear()
        top_words = self.actual_word_counts.most_common(100)

        filt = self.word_filter_edit.text().strip().lower()

        for word, count in top_words:
            if filt and filt not in word:
                continue
            pct = (count / self.total_word_count * 100) if self.total_word_count > 0 else 0
            item = QtWidgets.QTreeWidgetItem([word, f"{count:,}", f"{pct:.2f}"])
            item.setTextAlignment(1, QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            item.setTextAlignment(2, QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
            self.word_tree.addTopLevelItem(item)

    def _on_word_filter_change(self):
        if self.actual_word_counts:
            self._update_word_counts()

    def _show_word_context_menu(self, pos):
        item = self.word_tree.itemAt(pos)
        if item is None:
            return
        word = item.text(0).strip().lower()
        if not word:
            return
        menu = QtWidgets.QMenu(self.word_tree)
        act_stop = menu.addAction(_icon("fa5s.ban"), f"Add '{word}' to Stopwords")
        chosen = menu.exec(self.word_tree.viewport().mapToGlobal(pos))
        if chosen is act_stop:
            self.custom_stopwords.add(word)
            self._update_stopwords_display()
            self._generate_preview()

    # ================================================================
    # Generate / Preview
    # ================================================================

    def _set_rendering(self, active: bool) -> None:
        """Show/clear the in-progress indicator. Forces a paint so the label
        updates before the synchronous render call blocks the UI thread."""
        if not hasattr(self, "render_status"):
            return
        if active:
            self.render_status.setText("⏳ Rendering…")
            self.render_status.setStyleSheet(
                "font-size: 9pt; padding-left: 4px; font-weight: 600;"
            )
            QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.WaitCursor)
        else:
            self.render_status.setText("Live preview")
            self.render_status.setStyleSheet("font-size: 9pt; padding-left: 4px;")
            QtWidgets.QApplication.restoreOverrideCursor()
        QtWidgets.QApplication.processEvents()

    def _generate_preview(self):
        """Generate word cloud with current settings and update preview."""
        self._set_rendering(True)
        try:
            self._do_generate_preview()
        finally:
            self._set_rendering(False)

    def _do_generate_preview(self):
        colormap = WORDCLOUD_COLOR_SCHEMES.get(self.color_combo.currentText(), "Blues")
        background = WORDCLOUD_BACKGROUNDS.get(self.bg_combo.currentText(), "white")
        max_words = self.max_words_spin.value()

        # Shape mask
        shape_selection = self.shape_combo.currentText()
        if shape_selection == "Custom Image\u2026" and self.custom_mask_path:
            mask = load_custom_mask(self.custom_mask_path)
        else:
            shape_name = WORDCLOUD_SHAPES.get(shape_selection, "rectangle")
            mask = create_shape_mask(shape_name, width=1920, height=1080)

        # Stopwords (built-in English + custom)
        from sklearn.feature_extraction.text import ENGLISH_STOP_WORDS

        all_stopwords = set(ENGLISH_STOP_WORDS)
        all_stopwords.update(self.custom_stopwords)

        # Font size settings
        min_fs = self.min_font_spin.value()
        max_fs = self.max_font_spin.value()
        max_fs = max_fs if max_fs > 0 else None

        # Relative scaling
        rel_scaling = self.rel_scale_slider.value() / 100.0

        # Font path
        font_path = self._font_paths.get(self.font_combo.currentText())

        # Generate
        self.wc, self.word_frequencies = generate_wordcloud(
            self.texts,
            colormap=colormap,
            background_color=background,
            max_words=max_words,
            width=1920,
            height=1080,
            mask=mask,
            stopwords=all_stopwords,
            min_font_size=min_fs,
            max_font_size=max_fs,
            relative_scaling=rel_scaling,
            font_path=font_path,
        )

        if self.wc is None:
            QtWidgets.QMessageBox.critical(
                self,
                "Error",
                "Failed to generate word cloud. Is the wordcloud package installed?",
            )
            return

        # Compute actual word counts from original texts (excluding stopwords)
        all_stopwords_lower = {w.lower() for w in all_stopwords}
        all_words: list[str] = []
        for text in self.texts:
            words = re.findall(r"\b[a-zA-Z0-9]+\b", str(text).lower())
            words = [w for w in words if w not in all_stopwords_lower]
            all_words.extend(words)
        self.actual_word_counts = Counter(all_words)
        self.total_word_count = len(all_words)

        # Update stats bar
        unique_words = len(self.actual_word_counts)
        self.stats_label.setText(
            f"Unique words: {unique_words:,}  \u00b7  "
            f"Total words: {self.total_word_count:,}  \u00b7  "
            f"Texts: {len(self.texts):,}"
        )

        # Update word counts treeview
        self._update_word_counts()

        # Update preview image
        self._update_preview()

    def _update_preview(self):
        """Render the current word cloud into the preview label."""
        if self.wc is None:
            return

        img = wordcloud_to_image(self.wc)
        if img is None:
            return

        # Scale to fit the preview label while maintaining aspect ratio
        label_w = max(self.preview_label.width(), 200)
        label_h = max(self.preview_label.height(), 200)

        img_w, img_h = img.size
        scale_w = label_w / img_w
        scale_h = label_h / img_h
        self.preview_scale = min(scale_w, scale_h, 1.0)

        new_w = int(img_w * self.preview_scale)
        new_h = int(img_h * self.preview_scale)

        img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        qimage = ImageQt(img_resized)
        pixmap = QtGui.QPixmap.fromImage(qimage)
        self.preview_label.setPixmap(pixmap)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.wc is not None:
            self._update_preview()

    # ================================================================
    # Copy / Save
    # ================================================================

    def _copy_to_clipboard(self):
        if self.wc is None:
            QtWidgets.QMessageBox.warning(self, "No image", "Generate a word cloud first")
            return

        img = wordcloud_to_image(self.wc)
        if img is None:
            QtWidgets.QMessageBox.critical(self, "Error", "Failed to get image")
            return

        try:
            # Round-trip through PNG bytes so the QImage that lands on the
            # clipboard is fully self-contained (not a view over a buffer that
            # may be freed before paste). Works in PowerPoint, Word, browsers,
            # and image editors on both Windows and Linux.
            buf = io.BytesIO()
            if img.mode == "RGBA":
                # Flatten onto white so apps that don't honor alpha (PowerPoint)
                # don't paste a black background.
                background = Image.new("RGB", img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])
                background.save(buf, format="PNG")
            else:
                img.convert("RGB").save(buf, format="PNG")
            qimage = QtGui.QImage.fromData(buf.getvalue(), "PNG")
            if qimage.isNull():
                raise RuntimeError("Failed to decode rendered PNG")
            QtWidgets.QApplication.clipboard().setImage(qimage)
            QtWidgets.QMessageBox.information(
                self,
                "Copied",
                "Word cloud copied to clipboard!\nPaste with Ctrl+V (e.g. in PowerPoint).",
            )
        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self, "Error", f"Failed to copy to clipboard: {e}"
            )

    def _save_image(self):
        if self.wc is None:
            QtWidgets.QMessageBox.warning(self, "No image", "Generate a word cloud first")
            return

        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save Word Cloud",
            f"wordcloud_{self.column_name}.png",
            "PNG files (*.png);;JPEG files (*.jpg);;SVG files (*.svg);;All files (*)",
        )
        if not path:
            return

        if save_wordcloud(self.wc, path):
            QtWidgets.QMessageBox.information(
                self, "Saved", f"Word cloud saved to:\n{path}"
            )
        else:
            QtWidgets.QMessageBox.critical(self, "Error", "Failed to save word cloud")
