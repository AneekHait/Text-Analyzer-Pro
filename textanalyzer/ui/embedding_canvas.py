"""Inline matplotlib canvas widget for cluster scatter plots.

Matplotlib is the heaviest non-Qt dependency in this app (~500-700 ms cold
start on Windows), and the canvas is only used after the user clicks
*Visualize* on a clustered result. To keep the matplotlib load off the
launch path, the Figure + FigureCanvas are created lazily on the first
``render_embedding`` call instead of in ``__init__``. Existence of the
matplotlib package is checked via ``importlib.util.find_spec`` — that
returns the spec without actually executing the module.
"""

import importlib.util as _ilu

import numpy as np
from PySide6 import QtCore, QtWidgets

_HAS_MPL = _ilu.find_spec("matplotlib") is not None


class EmbeddingCanvas(QtWidgets.QWidget):
    """Embedded matplotlib canvas for cluster scatter plots; replaces blocking plt.show()."""

    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self._layout = layout
        self._canvas = None
        self._figure = None
        self._placeholder = QtWidgets.QLabel("Run clustering and click Visualize to see a 2D projection.")
        self._placeholder.setAlignment(QtCore.Qt.AlignCenter)
        self._placeholder.setProperty("role", "muted")
        self._placeholder.setMinimumHeight(220)
        layout.addWidget(self._placeholder)
        # Matplotlib Figure + FigureCanvas created on first render
        # (see _ensure_canvas) — keeps app startup off the matplotlib path.

    def _ensure_canvas(self) -> bool:
        """Lazily create the matplotlib canvas. Returns True if it's now available."""
        if self._canvas is not None:
            return True
        if not _HAS_MPL:
            return False
        try:
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
        except Exception:
            return False
        self._figure = Figure(figsize=(6, 4), tight_layout=True)
        self._canvas = FigureCanvas(self._figure)
        self._canvas.hide()
        self._layout.addWidget(self._canvas, 1)
        return True

    def is_available(self) -> bool:
        """Whether matplotlib is installed (canvas can be created on demand)."""
        return _HAS_MPL

    def clear(self):
        if self._figure is not None:
            self._figure.clear()
            if self._canvas is not None:
                self._canvas.hide()
        self._placeholder.show()

    def render_embedding(self, embedding_2d: np.ndarray, labels: np.ndarray, method: str, bg: str, fg: str, muted: str):
        if not self._ensure_canvas():
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
