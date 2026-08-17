"""Categorization configuration dialog."""

from PySide6 import QtCore, QtWidgets

from textanalyzer.models.config import CategorizationConfig


class CategorizationDialog(QtWidgets.QDialog):
    """Modal for configuring and launching the Run Categorization workflow."""

    def __init__(self, settings: dict, preferred_vectorizer: str = "embedding", parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Run Categorization")
        self.setModal(True)
        self._build_ui(preferred_vectorizer)

    def _build_ui(self, preferred_vectorizer: str) -> None:
        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(10)

        form = QtWidgets.QFormLayout()
        form.setHorizontalSpacing(10)
        form.setVerticalSpacing(10)

        last_granularity = int(self.settings.get("last_granularity", 50) or 50)
        self.slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.slider.setRange(0, 100)
        self.slider.setValue(last_granularity)
        self.slider.setToolTip("Coarse (left) → Fine (right). Drives HDBSCAN min_cluster_size.")
        self.live_label = QtWidgets.QLabel()
        self.live_label.setProperty("role", "muted")

        self.vec_combo = QtWidgets.QComboBox()
        self.vec_combo.addItem("Embeddings (semantic)", userData="embedding")
        self.vec_combo.addItem("TF-IDF (lexical)", userData="tfidf")
        idx = 0 if preferred_vectorizer == "embedding" else 1
        self.vec_combo.setCurrentIndex(idx)

        form.addRow("Granularity:", self.slider)
        form.addRow("", self.live_label)
        form.addRow("Vectorizer:", self.vec_combo)

        # Advanced expander
        adv_toggle = QtWidgets.QToolButton()
        adv_toggle.setText("Advanced (raw HDBSCAN knobs)")
        adv_toggle.setCheckable(True)
        adv_toggle.setArrowType(QtCore.Qt.RightArrow)
        adv_toggle.setStyleSheet("QToolButton { border: none; }")
        self.adv_widget = QtWidgets.QWidget()
        adv_form = QtWidgets.QFormLayout(self.adv_widget)
        adv_form.setContentsMargins(20, 0, 0, 0)
        adv_form.setHorizontalSpacing(10)
        adv_form.setVerticalSpacing(8)

        suggested_mcs = CategorizationConfig.min_cluster_size_from_granularity(last_granularity)
        self.mcs_spin = QtWidgets.QSpinBox()
        self.mcs_spin.setRange(2, 200)
        self.mcs_spin.setValue(
            int(self.settings.get("last_min_cluster_size", suggested_mcs) or suggested_mcs)
        )
        adv_form.addRow("Min sub-cluster size:", self.mcs_spin)

        self.nr_spin = QtWidgets.QSpinBox()
        self.nr_spin.setRange(1, 500)
        self.nr_spin.setValue(
            int(self.settings.get("last_non_repetitive_min_size", 5) or 5)
        )
        adv_form.addRow("Non-Repetitive cutoff:", self.nr_spin)

        self.adv_widget.setVisible(False)

        def _toggle_adv(checked: bool) -> None:
            self.adv_widget.setVisible(checked)
            adv_toggle.setArrowType(
                QtCore.Qt.DownArrow if checked else QtCore.Qt.RightArrow
            )

        adv_toggle.toggled.connect(_toggle_adv)
        self.slider.valueChanged.connect(self._update_label)
        self._update_label()

        layout.addLayout(form)
        layout.addWidget(adv_toggle)
        layout.addWidget(self.adv_widget)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttons.button(QtWidgets.QDialogButtonBox.Ok).setText("Run")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _update_label(self) -> None:
        g = self.slider.value()
        suggested = CategorizationConfig.min_cluster_size_from_granularity(g)
        self.live_label.setText(f"min_cluster_size ≈ {suggested}  (slider value: {g})")
        if not self.adv_widget.isVisible():
            self.mcs_spin.blockSignals(True)
            self.mcs_spin.setValue(suggested)
            self.mcs_spin.blockSignals(False)

    def get_categorize_kwargs(self) -> dict:
        """Return the kwargs dict ready for controller.run_categorization."""
        return {
            "vectorizer_kind": str(self.vec_combo.currentData() or "embedding"),
            "min_cluster_size": int(self.mcs_spin.value()),
            "non_repetitive_min_size": int(self.nr_spin.value()),
        }

    def granularity(self) -> int:
        return int(self.slider.value())

    def min_cluster_size(self) -> int:
        return int(self.mcs_spin.value())

    def non_repetitive_min_size(self) -> int:
        return int(self.nr_spin.value())
