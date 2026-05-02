"""Analysis service — cleaning preview & result-model construction."""

from __future__ import annotations

from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd

from textanalyzer.engine.cluster import TextCleaningResult, coerce_text_column, prepare_text_cleaning

from ..models.config import CleaningConfigModel
from ..models.result import ClusterResultModel


class AnalysisService:
    """Stateless helpers that bridge raw worker output to domain models."""

    @staticmethod
    def run_cleaning_preview(
        texts: List[str], config: CleaningConfigModel
    ) -> TextCleaningResult:
        """Run ``prepare_text_cleaning`` and return the engine result."""
        return prepare_text_cleaning(texts, config.to_engine_config())

    @staticmethod
    def build_result_model(
        worker_result: dict,
        df: pd.DataFrame,
        source_column: str,
        cleaned_column_name: str,
    ) -> ClusterResultModel:
        """Convert the raw dict emitted by ``ClusterWorker.finished`` into a
        :class:`ClusterResultModel` and apply side-effects on *df*
        (adds label + cleaned columns).
        """
        cleaning_result: TextCleaningResult = worker_result["cleaning_result"]

        # Write cleaned texts into df.
        df[cleaned_column_name] = cleaning_result.cleaned_texts

        # Expand per-representative labels to full-df labels.
        labels_arr = np.asarray(worker_result["labels"], dtype=int)
        label_by_rep = {
            rep_idx: int(lbl)
            for rep_idx, lbl in zip(cleaning_result.kept_indices, labels_arr)
        }
        full_labels = np.full(len(df), -1, dtype=int)
        for row_idx, rep_idx in enumerate(cleaning_result.representative_index_by_row):
            if rep_idx is not None:
                full_labels[row_idx] = label_by_rep[rep_idx]
        df["cluster_label"] = full_labels

        X = worker_result["X"]
        return ClusterResultModel(
            labels=full_labels,
            kept_labels=labels_arr,
            cluster_names=worker_result["cluster_names"],
            top_keywords=worker_result["top_keywords"],
            X=X,
            vectorizer=worker_result["vectorizer"],
            model=worker_result["model"],
            cleaned_column_name=cleaned_column_name,
            cleaning_stats=cleaning_result.stats,
            n_documents=X.shape[0] if X is not None else 0,
            n_features=X.shape[1] if X is not None else 0,
        )

    @staticmethod
    def coerce_column(series: pd.Series) -> pd.Series:
        return coerce_text_column(series)


__all__ = ["AnalysisService"]
