"""Inline matplotlib canvas widget for cluster scatter plots."""

import numpy as np
from PySide6 import QtCore, QtWidgets

try:
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
    _HAS_MPL = True
except Exception:  # pragma: no cover
    Figure = None
    FigureCanvas = None
    _HAS_MPL = False


class EmbeddingCanvas(QtWidgets.QWidget):
    """Embedded matplotlib canvas for cluster scatter plots; replaces blocking plt.show()."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self._canvas = None
        self._figure = None
        self._placeholder = QtWidgets.QLabel("Run clustering and click Visualize to see a 2D projection.")
        self._placeholder.setAlignment(QtCore.Qt.AlignCenter)
        self._placeholder.setProperty("role", "muted")
        self._placeholder.setMinimumHeight(220)
        layout.addWidget(self._placeholder)

        if _HAS_MPL:
            self._figure = Figure(figsize=(6, 4), tight_layout=True)
            self._canvas = FigureCanvas(self._figure)
            self._canvas.hide()
            layout.addWidget(self._canvas, 1)

    def is_available(self) -> bool:
        return _HAS_MPL and self._canvas is not None

    def clear(self):
        if self._figure is not None:
            self._figure.clear()
            if self._canvas is not None:
                self._canvas.hide()
        self._placeholder.show()

    def render_embedding(self, embedding_2d: np.ndarray, labels: np.ndarray, method: str, bg: str, fg: str, muted: str):
        if not self.is_available():
            return
        self._placeholder.hide()
        self._figure.clear()
        self._figure.patch.set_facecolor(bg)
        ax = self._figure.add_subplot(111)
        ax.set_facecolor(bg)
        unique_labels = sorted({int(l) for l in labels})
        cmap = ["#3d6dff", "#1f9d55", "#d29922", "#d73a49", "#6f4dff", "#19b6c0", "#e0529c", "#7a8493", "#5b8dff", "#2ea864"]
        for i, label in enumerate(unique_labels):
            mask = labels == label
            if not np.any(mask):
                continue
            color = "#7a8493" if label == -1 else cmap[i % len(cmap)]
            ax.scatter(embedding_2d[mask, 0], embedding_2d[mask, 1], s=14, alpha=0.75,
                       color=color, label=str(label), edgecolors="none")
        ax.set_title(f"Cluster visualization ({method.upper()})", color=fg, fontsize=11)
        ax.tick_params(colors=muted, labelsize=8)
        for spine in ax.spines.values():
            spine.set_color(muted)
        legend = ax.legend(loc="best", fontsize=8, frameon=False, labelcolor=fg, title="Cluster")
        if legend is not None and legend.get_title() is not None:
            legend.get_title().set_color(fg)
        self._canvas.show()
        self._canvas.draw_idle()
