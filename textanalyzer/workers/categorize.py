"""Background categorization worker (runs on a QThread).

Discovers a single-level subcategory taxonomy from the active dataframe:
cleans → embeds → HDBSCAN → marks Non-Repetitive → names subclusters via
deterministic Title Case phrasing. Optionally applies a pre-trained taxonomy
when ``taxonomy_payload`` is supplied (fast path, skips HDBSCAN entirely).
"""

from PySide6 import QtCore

from textanalyzer.engine.cluster import (
    apply_taxonomy,
    categorize_taxonomy,
    prepare_text_cleaning,
)


class CategorizeWorker(QtCore.QObject):
    progress = QtCore.Signal(int, str)
    finished = QtCore.Signal(object)
    failed = QtCore.Signal(str)

    def __init__(
        self,
        df,
        column,
        cleaning_config,
        categorize_kwargs=None,
        user_renames=None,
        taxonomy_payload=None,
        confidence_threshold=0.45,
    ):
        super().__init__()
        self.df = df
        self.column = column
        self.cleaning_config = cleaning_config
        self.categorize_kwargs = dict(categorize_kwargs or {})
        self.user_renames = dict(user_renames or {})
        self.taxonomy_payload = taxonomy_payload
        self.confidence_threshold = float(confidence_threshold)
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def _check_cancel(self):
        if self._cancelled:
            raise RuntimeError("Cancelled by user.")

    @QtCore.Slot()
    def run(self):
        try:
            self.progress.emit(3, "Cleaning text…")
            cleaning_result = prepare_text_cleaning(
                self.df[self.column].tolist(), self.cleaning_config
            )
            self._check_cancel()
            if not cleaning_result.cluster_input_texts:
                raise ValueError(
                    "Cleaning produced no usable text rows. "
                    "Adjust the cleaning settings and try again."
                )

            if self.taxonomy_payload is not None:
                self.progress.emit(40, "Applying saved taxonomy…")
                result = apply_taxonomy(
                    cleaning_result.cluster_input_texts,
                    self.taxonomy_payload,
                    confidence_threshold=self.confidence_threshold,
                )
                self.progress.emit(100, "Done")
            else:
                def _progress_cb(pct: int, msg: str) -> None:
                    # Engine emits 5/40/70/90/100 — pass through verbatim so
                    # the action-bar label matches the engine's stage names.
                    self.progress.emit(int(pct), str(msg))
                    self._check_cancel()

                result = categorize_taxonomy(
                    cleaning_result.cluster_input_texts,
                    user_renames=self.user_renames,
                    progress_cb=_progress_cb,
                    **self.categorize_kwargs,
                )

            # Stash the encoded matrix the engine used so the controller can
            # offer merge_clusters / split_cluster post-hoc without re-running.
            X_used = result.vectorizer.transform(cleaning_result.cluster_input_texts) \
                if hasattr(result.vectorizer, "transform") else None

            self.finished.emit({
                "cleaning_result": cleaning_result,
                "taxonomy_result": result,
                "applied_from_saved": self.taxonomy_payload is not None,
                "X": X_used,
                "texts": list(cleaning_result.cluster_input_texts),
            })
        except Exception as error:
            self.failed.emit(str(error))
