"""Analysis controller — owns the session, workers, and mediates signals.

The controller is a ``QObject`` with rich signals that views subscribe to.
It owns the :class:`AnalysisSession` and is the *only* object that talks
to :class:`AnalysisService`, :class:`IOService`, and workers.
"""

from __future__ import annotations

import os
from typing import Optional

import numpy as np
from PySide6 import QtCore, QtWidgets

from textanalyzer import settings as app_settings
from textanalyzer.models.config import CleaningConfigModel
from textanalyzer.models.result import ClusterResultModel
from textanalyzer.models.session import AnalysisSession
from textanalyzer.services.analysis import AnalysisService
from textanalyzer.services.io import IOService
from textanalyzer.workers.categorize import CategorizeWorker
from textanalyzer.workers.cluster import ClusterWorker
from textanalyzer.workers.embedding import EmbeddingWorker


class AnalysisController(QtCore.QObject):
    """Mediator between AnalysisTab views and backend services."""

    # -- Signals for views -------------------------------------------
    file_loaded = QtCore.Signal(str, list, str)  # path, sheet_names, default_output
    sheet_loaded = QtCore.Signal(list, str)       # column_names, restored_column
    cleaning_preview_ready = QtCore.Signal(object)  # TextCleaningResult
    cleaning_preview_failed = QtCore.Signal(str)
    cluster_started = QtCore.Signal()
    cluster_progress = QtCore.Signal(int, str)     # value, message
    cluster_finished = QtCore.Signal(object)       # ClusterResultModel
    cluster_failed = QtCore.Signal(str)
    cluster_cancelled = QtCore.Signal()
    categorization_started = QtCore.Signal()
    categorization_progress = QtCore.Signal(int, str)  # value, message
    categorization_finished = QtCore.Signal(object)    # TaxonomyResult
    categorization_failed = QtCore.Signal(str)
    categorization_cancelled = QtCore.Signal()
    taxonomy_saved = QtCore.Signal(str)
    taxonomy_loaded = QtCore.Signal(str)
    embedding_ready = QtCore.Signal(object, object, object)  # coords, labels, names
    embedding_failed = QtCore.Signal(str)
    results_saved = QtCore.Signal(str)
    model_saved = QtCore.Signal(str)
    log = QtCore.Signal(str)

    def __init__(self, settings: dict, parent: QtCore.QObject | None = None) -> None:
        super().__init__(parent)
        self.session = AnalysisSession()
        self.settings = settings
        self._cluster_thread: QtCore.QThread | None = None
        self._cluster_worker: ClusterWorker | None = None
        self._embedding_thread: QtCore.QThread | None = None
        self._embedding_worker: EmbeddingWorker | None = None
        self._categorize_thread: QtCore.QThread | None = None
        self._categorize_worker: CategorizeWorker | None = None
        # Persistent across sessions: fingerprint -> user-edited subcategory name.
        self.user_taxonomy_renames: dict[str, str] = dict(
            self.settings.get("user_taxonomy_renames", {}) or {}
        )
        # In-memory only: the last TaxonomyResult and where it came from.
        self.taxonomy_result = None
        self.taxonomy_cleaning_result = None
        self.taxonomy_applied_from_saved = False
        self.taxonomy_X = None
        self.taxonomy_texts: list[str] = []

    # ----------------------------------------------------------------
    # File loading
    # ----------------------------------------------------------------
    def open_file(self, path: str) -> None:
        if not path or not os.path.isfile(path):
            return
        try:
            sheets = IOService.sheet_names(path)
            ext = IOService.file_extension(path)
            self.session.file_path = path
            self.session.output_path = IOService.default_output_path(path)
            self.session.reset_results()

            # Recent-file tracking.
            app_settings.add_recent_file(self.settings, path)
            app_settings.save(self.settings)

            self.file_loaded.emit(path, sheets, self.session.output_path)
            self.log.emit(f"Opened: {os.path.basename(path)} ({len(sheets)} sheet(s))")

            if sheets:
                self.load_sheet(sheets[0])
        except Exception as exc:
            self.log.emit(f"\u2717 Error reading file: {exc}")

    def load_sheet(self, sheet_name: str) -> None:
        if not self.session.file_path:
            return
        try:
            ext = IOService.file_extension(self.session.file_path)
            from textanalyzer.engine.cluster import EXCEL_INPUT_EXTENSIONS
            selected = sheet_name if ext in EXCEL_INPUT_EXTENSIONS else None
            self.session.df = IOService.load_table(self.session.file_path, sheet_name=selected)
            self.session.sheet_name = sheet_name
            cols = list(self.session.df.columns)
            restored = self.settings.get("last_column", "")
            self.sheet_loaded.emit(cols, restored)
            self.log.emit(f"Sheet '{sheet_name}': {len(self.session.df)} rows, {len(cols)} columns")
        except Exception as exc:
            self.log.emit(f"\u2717 Error loading sheet: {exc}")

    # ----------------------------------------------------------------
    # Cleaning
    # ----------------------------------------------------------------
    def set_cleaning_config(self, config: CleaningConfigModel) -> None:
        self.session.cleaning_config = config

    def refresh_cleaning_preview(self, column: str, config: CleaningConfigModel) -> None:
        if self.session.df is None or not column:
            return
        try:
            texts = self.session.df[column].tolist()
            result = AnalysisService.run_cleaning_preview(texts, config)
            self.cleaning_preview_ready.emit(result)
        except Exception as exc:
            self.cleaning_preview_failed.emit(str(exc))
            self.log.emit(f"\u2717 Cleaning preview failed: {exc}")

    # ----------------------------------------------------------------
    # Clustering
    # ----------------------------------------------------------------
    def run_clustering(
        self,
        column: str,
        config: CleaningConfigModel,
        algorithm: str,
        n_clusters: int,
        top_n: int,
        joiner: str,
        vectorize_kwargs: dict | None = None,
        cluster_kwargs: dict | None = None,
    ) -> None:
        if self.session.df is None:
            self.cluster_failed.emit("No data loaded")
            return
        if self._cluster_thread is not None and self._cluster_thread.isRunning():
            return

        self.session.column = column
        self.session.cleaning_config = config

        worker = ClusterWorker(
            df=self.session.df,
            column=column,
            cleaning_config=config.to_engine_config(),
            algorithm=algorithm,
            n_clusters=n_clusters,
            top_n=top_n,
            joiner=joiner,
            vectorize_kwargs=vectorize_kwargs,
            cluster_kwargs=cluster_kwargs,
        )
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self.cluster_progress.emit)
        worker.finished.connect(self._on_cluster_finished)
        worker.failed.connect(self._on_cluster_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._cluster_thread = thread
        self._cluster_worker = worker

        self.cluster_started.emit()
        self.log.emit("=" * 60)
        self.log.emit(f"Starting clustering (Algorithm: {algorithm}, n_clusters: {n_clusters})")
        thread.start()

    def cancel_clustering(self) -> None:
        if self._cluster_worker is not None:
            self._cluster_worker.cancel()
            self.log.emit("Cancellation requested\u2026")

    def _on_cluster_finished(self, raw_result: dict) -> None:
        try:
            col = self.session.column or ""
            cleaned_col = f"{col}_cleaned"
            model = AnalysisService.build_result_model(
                raw_result, self.session.df, col, cleaned_col
            )
            self.session.result = model
            self.session.cleaning_result = raw_result["cleaning_result"]
            self.session.user_cluster_names = dict(model.cluster_names)
            self.session.mark_dirty()

            self.log.emit(f"  Vectorizer: {model.n_documents} documents, {model.n_features} features")
            self.log.emit("\u2713 Clustering finished. Edit names below and click Save Results.")

            self.settings["last_column"] = col
            app_settings.save(self.settings)

            self.cluster_finished.emit(model)
        finally:
            self._cluster_thread = None
            self._cluster_worker = None

    def _on_cluster_failed(self, error_message: str) -> None:
        self.log.emit(f"\u2717 Clustering error: {error_message}")
        if "Cancelled by user" in error_message:
            self.cluster_cancelled.emit()
        else:
            self.cluster_failed.emit(error_message)
        self._cluster_thread = None
        self._cluster_worker = None

    # ----------------------------------------------------------------
    # Categorization (single-level taxonomy discovery)
    # ----------------------------------------------------------------
    def run_categorization(
        self,
        column: str,
        config: CleaningConfigModel,
        categorize_kwargs: dict | None = None,
        taxonomy_payload: dict | None = None,
        confidence_threshold: float = 0.45,
    ) -> None:
        if self.session.df is None:
            self.categorization_failed.emit("No data loaded")
            return
        if self._categorize_thread is not None and self._categorize_thread.isRunning():
            return

        self.session.column = column
        self.session.cleaning_config = config

        worker = CategorizeWorker(
            df=self.session.df,
            column=column,
            cleaning_config=config.to_engine_config(),
            categorize_kwargs=categorize_kwargs,
            user_renames=self.user_taxonomy_renames,
            taxonomy_payload=taxonomy_payload,
            confidence_threshold=confidence_threshold,
        )
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress.connect(self.categorization_progress.emit)
        worker.finished.connect(self._on_categorization_finished)
        worker.failed.connect(self._on_categorization_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._categorize_thread = thread
        self._categorize_worker = worker

        self.categorization_started.emit()
        self.log.emit("=" * 60)
        if taxonomy_payload is not None:
            self.log.emit("Applying saved taxonomy (fast path)…")
        else:
            self.log.emit("Starting categorization (HDBSCAN + keyword phrasing)…")
        thread.start()

    def cancel_categorization(self) -> None:
        if self._categorize_worker is not None:
            self._categorize_worker.cancel()
            self.log.emit("Cancellation requested…")

    def _on_categorization_finished(self, raw: dict) -> None:
        try:
            result = raw["taxonomy_result"]
            self.taxonomy_result = result
            self.taxonomy_cleaning_result = raw.get("cleaning_result")
            self.taxonomy_applied_from_saved = bool(raw.get("applied_from_saved"))
            # Stash matrix + texts so the GUI can call merge_clusters /
            # split_cluster without re-running the worker.
            self.taxonomy_X = raw.get("X")
            self.taxonomy_texts = raw.get("texts") or []
            stats = result.stats or {}
            self.log.emit(
                f"  Subcategories: {stats.get('n_subclusters', 0)}, "
                f"Non-Repetitive: {stats.get('pct_non_repetitive', 0.0):.0%}"
            )
            self.log.emit("✓ Categorization finished. Review the Subcategory column and click Save Results.")
            self.settings["last_column"] = self.session.column or ""
            app_settings.save(self.settings)
            self.categorization_finished.emit(result)
        finally:
            self._categorize_thread = None
            self._categorize_worker = None

    def _on_categorization_failed(self, error_message: str) -> None:
        self.log.emit(f"✗ Categorization error: {error_message}")
        if "Cancelled by user" in error_message:
            self.categorization_cancelled.emit()
        else:
            self.categorization_failed.emit(error_message)
        self._categorize_thread = None
        self._categorize_worker = None

    def record_taxonomy_rename(self, fingerprint: str, new_name: str) -> None:
        """Persist a user rename so re-runs (and Save Taxonomy) keep it."""
        if not fingerprint:
            return
        cleaned = (new_name or "").strip()
        if cleaned:
            self.user_taxonomy_renames[fingerprint] = cleaned
        else:
            self.user_taxonomy_renames.pop(fingerprint, None)
        self.settings["user_taxonomy_renames"] = dict(self.user_taxonomy_renames)
        app_settings.save(self.settings)

    def save_taxonomy(self, path: str) -> None:
        if self.taxonomy_result is None:
            self.categorization_failed.emit("Nothing to save — run categorization first")
            return
        try:
            IOService.save_taxonomy(
                self.taxonomy_result, path,
                user_renames=self.user_taxonomy_renames,
            )
            self.log.emit(f"✓ Taxonomy saved to {path}")
            self.taxonomy_saved.emit(path)
        except Exception as exc:
            self.log.emit(f"✗ Taxonomy save failed: {exc}")
            self.categorization_failed.emit(str(exc))

    def edit_taxonomy(
        self,
        new_result,
        cleaning_result=None,
    ) -> None:
        """Adopt a rebuilt TaxonomyResult (from merge_clusters / split_cluster).

        Wired through ``categorization_finished`` so the Results-tab refresh
        path is the same as a fresh run. Cleaning result is reused unchanged.
        """
        self.taxonomy_result = new_result
        if cleaning_result is not None:
            self.taxonomy_cleaning_result = cleaning_result
        # Re-emit finished so the GUI re-populates the cluster cards.
        self.categorization_finished.emit(new_result)

    def load_taxonomy(self, path: str) -> dict | None:
        """Load a saved taxonomy payload. Caller runs categorization with it."""
        try:
            payload = IOService.load_taxonomy(path)
            n_subs = len(payload.get("subcategory_names") or {})
            self.log.emit(f"✓ Loaded taxonomy: {n_subs} subcategories")
            # Surface the manifest summary so the user can audit provenance
            # before applying to fresh data.
            try:
                self.log.emit("  " + IOService.describe_taxonomy(payload))
            except Exception:
                pass
            # Merge saved user_renames into the in-memory map so subsequent
            # runs (even ones that discover fresh clusters) honor the edits.
            saved_renames = payload.get("user_renames") or {}
            if saved_renames:
                self.user_taxonomy_renames.update(saved_renames)
                self.settings["user_taxonomy_renames"] = dict(self.user_taxonomy_renames)
                app_settings.save(self.settings)
            self.taxonomy_loaded.emit(path)
            return payload
        except Exception as exc:
            self.log.emit(f"✗ Taxonomy load failed: {exc}")
            self.categorization_failed.emit(str(exc))
            return None

    # ----------------------------------------------------------------
    # Visualization
    # ----------------------------------------------------------------
    def visualize(self, method: str = "pca") -> None:
        result = self.session.result
        if result is None or result.X is None:
            self.embedding_failed.emit("Run clustering first")
            return
        if self._embedding_thread is not None and self._embedding_thread.isRunning():
            return
        self.log.emit(f"Generating {method.upper()} visualization\u2026")
        worker = EmbeddingWorker(result.X, method)
        thread = QtCore.QThread(self)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_embedding_ready)
        worker.failed.connect(self._on_embedding_failed)
        worker.finished.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._embedding_thread = thread
        self._embedding_worker = worker
        thread.start()

    def _on_embedding_ready(self, coords) -> None:
        result = self.session.result
        if result is None:
            return
        self.log.emit(f"\u2713 Visualization ready")
        self.embedding_ready.emit(coords, result.kept_labels, result.cluster_names)
        self._embedding_thread = None
        self._embedding_worker = None

    def _on_embedding_failed(self, msg: str) -> None:
        self.log.emit(f"\u2717 Visualization failed: {msg}")
        self.embedding_failed.emit(msg)
        self._embedding_thread = None
        self._embedding_worker = None

    # ----------------------------------------------------------------
    # Save / export
    # ----------------------------------------------------------------
    def save_results(self, user_names: dict[int, str], output_path: str) -> None:
        df = self.session.df
        result = self.session.result
        if df is None or result is None:
            self.cluster_failed.emit("Nothing to save — run clustering first")
            return
        try:
            df["cluster_name"] = [
                user_names.get(int(lbl), "") for lbl in result.labels
            ]
            saved = IOService.save_results(df, output_path)
            self.session.mark_clean()
            self.log.emit(f"\u2713 Results saved to {saved}")
            self.results_saved.emit(saved)
        except Exception as exc:
            self.log.emit(f"\u2717 Save failed: {exc}")
            self.cluster_failed.emit(str(exc))

    def save_model(self, path: str) -> None:
        result = self.session.result
        if result is None or result.model is None:
            self.cluster_failed.emit("Nothing to save — run clustering first")
            return
        try:
            IOService.save_model(
                result.model,
                result.vectorizer,
                self.session.user_cluster_names,
                result.top_keywords,
                path,
            )
            self.log.emit(f"\u2713 Model saved to {path}")
            self.model_saved.emit(path)
        except Exception as exc:
            self.log.emit(f"\u2717 Model save failed: {exc}")
            self.cluster_failed.emit(str(exc))


__all__ = ["AnalysisController"]
