"""Background worker for text cleaning + vectorization (Suggest K / Compare)."""

from PySide6 import QtCore

from textanalyzer.engine.cluster import prepare_text_cleaning, vectorize_texts


class PrepareCorpusWorker(QtCore.QObject):
    """Cleans and vectorizes a text column off the main thread."""

    finished = QtCore.Signal(object, object)  # (vectorizer, X)
    failed = QtCore.Signal(str)

    def __init__(self, texts, cleaning_config, vectorize_kwargs=None):
        super().__init__()
        self.texts = texts
        self.cleaning_config = cleaning_config
        self.vectorize_kwargs = dict(vectorize_kwargs or {})

    @QtCore.Slot()
    def run(self):
        try:
            cleaning_result = prepare_text_cleaning(self.texts, self.cleaning_config)
            if not cleaning_result.cluster_input_texts:
                self.failed.emit(
                    "Cleaning produced no usable rows. Adjust cleaning settings and retry."
                )
                return
            vectorizer, X = vectorize_texts(
                cleaning_result.cluster_input_texts, **self.vectorize_kwargs
            )
            self.finished.emit(vectorizer, X)
        except Exception as exc:
            self.failed.emit(str(exc))
