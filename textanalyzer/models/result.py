"""Cluster-result model — lightweight value object holding output of a run."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import numpy as np


@dataclass
class ClusterResultModel:
    """Immutable snapshot of one clustering run.

    Created by the analysis controller after the :class:`ClusterWorker`
    emits its ``finished`` signal. Widgets can read from this model
    without touching workers or engine APIs directly.
    """

    labels: np.ndarray                                     # full-length (one per df row), -1 for excluded
    kept_labels: np.ndarray                                # labels for representative rows
    cluster_names: Dict[int, str]                          # {cid: auto-name}
    top_keywords: Dict[int, List[tuple]]                   # {cid: [(term, score), ...]}
    X: Any = None                                          # sparse feature matrix
    vectorizer: Any = None                                 # fitted TfidfVectorizer
    model: Any = None                                      # fitted clustering model
    cleaned_column_name: str = ""
    cleaning_stats: Dict[str, Any] = field(default_factory=dict)
    n_documents: int = 0
    n_features: int = 0

    @property
    def n_clusters(self) -> int:
        return len(self.cluster_names)

    def row_count_for(self, cid: int) -> int:
        return int(np.sum(self.labels == cid))


__all__ = ["ClusterResultModel"]
