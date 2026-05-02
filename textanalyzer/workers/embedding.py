"""Background 2D embedding worker (PCA / t-SNE)."""

from PySide6 import QtCore

from cluster_tool import compute_embedding_2d


class EmbeddingWorker(QtCore.QObject):
    finished = QtCore.Signal(object, str)
    failed = QtCore.Signal(str)

    def __init__(self, X, method):
        super().__init__()
        self.X = X
        self.method = method

    @QtCore.Slot()
    def run(self):
        try:
            emb = compute_embedding_2d(self.X, method=self.method)
            self.finished.emit(emb, self.method)
        except Exception as error:
            self.failed.emit(str(error))
