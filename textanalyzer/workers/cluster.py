"""Background clustering worker (runs on a QThread)."""

from PySide6 import QtCore

from textanalyzer.engine.cluster import (
    assign_cluster_names,
    cluster_texts,
    get_top_keywords_per_cluster,
    prepare_text_cleaning,
    vectorize_texts,
)


class ClusterWorker(QtCore.QObject):
    progress = QtCore.Signal(int, str)
    finished = QtCore.Signal(object)
    failed = QtCore.Signal(str)

    def __init__(
        self,
        df,
        column,
        cleaning_config,
        algorithm,
        n_clusters,
        top_n,
        joiner,
        vectorize_kwargs=None,
        cluster_kwargs=None,
    ):
        super().__init__()
        self.df = df
        self.column = column
        self.cleaning_config = cleaning_config
        self.algorithm = algorithm
        self.n_clusters = n_clusters
        self.top_n = top_n
        self.joiner = joiner
        self.vectorize_kwargs = dict(vectorize_kwargs or {})
        self.cluster_kwargs = dict(cluster_kwargs or {})
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def _check_cancel(self):
        if self._cancelled:
            raise RuntimeError("Cancelled by user.")

    @QtCore.Slot()
    def run(self):
        try:
            self.progress.emit(5, "Cleaning text…")
            cleaning_result = prepare_text_cleaning(self.df[self.column].tolist(), self.cleaning_config)
            self._check_cancel()
            if not cleaning_result.cluster_input_texts:
                raise ValueError("Cleaning produced no usable text rows. Adjust the cleaning settings and try again.")
            kept = cleaning_result.stats["kept_row_count"]
            if self.algorithm != "dbscan" and self.n_clusters > kept:
                raise ValueError(
                    f"n_clusters ({self.n_clusters}) cannot exceed the cleaned clustering row count ({kept})."
                )

            vec_kind = self.vectorize_kwargs.get("vectorizer_kind", "tfidf")
            if vec_kind == "embedding":
                self.progress.emit(35, "Generating embeddings… (first run may download a model)")
            else:
                self.progress.emit(35, "Vectorizing…")
            vectorizer, X = vectorize_texts(
                cleaning_result.cluster_input_texts, **self.vectorize_kwargs
            )
            self._check_cancel()

            self.progress.emit(65, f"Clustering ({self.algorithm})…")
            model, labels = cluster_texts(
                X,
                algorithm=self.algorithm,
                n_clusters=self.n_clusters,
                **self.cluster_kwargs,
            )
            self._check_cancel()

            self.progress.emit(85, "Naming clusters…")
            top_keywords = get_top_keywords_per_cluster(vectorizer, X, labels, top_n=10)
            cluster_names = assign_cluster_names(top_keywords, name_top_n=self.top_n, joiner=self.joiner)

            self.progress.emit(100, "Done")
            self.finished.emit(
                {
                    "cleaning_result": cleaning_result,
                    "vectorizer": vectorizer,
                    "X": X,
                    "model": model,
                    "labels": labels,
                    "top_keywords": top_keywords,
                    "cluster_names": cluster_names,
                }
            )
        except Exception as error:
            self.failed.emit(str(error))
