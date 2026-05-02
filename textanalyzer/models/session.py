"""Analysis session — holds all mutable state for one analysis tab."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd

from .config import CleaningConfigModel
from .result import ClusterResultModel


@dataclass
class AnalysisSession:
    """Represents the full state of one analysis workflow.

    Each workspace tab owns one ``AnalysisSession``.  The controller
    mutates it in response to user actions; widgets observe it through
    signals on the controller.
    """

    # Source ----------------------------------------------------------
    file_path: Optional[str] = None
    sheet_name: Optional[str] = None
    column: Optional[str] = None
    df: Optional[pd.DataFrame] = None

    # Cleaning --------------------------------------------------------
    cleaning_config: CleaningConfigModel = field(default_factory=CleaningConfigModel)
    cleaning_result: Any = None          # TextCleaningResult from engine

    # Clustering ------------------------------------------------------
    result: Optional[ClusterResultModel] = None
    user_cluster_names: Dict[int, str] = field(default_factory=dict)

    # I/O -------------------------------------------------------------
    output_path: str = ""

    # Dirty tracking --------------------------------------------------
    _dirty: bool = False

    @property
    def is_dirty(self) -> bool:
        return self._dirty

    def mark_dirty(self) -> None:
        self._dirty = True

    def mark_clean(self) -> None:
        self._dirty = False

    @property
    def has_data(self) -> bool:
        return self.df is not None

    @property
    def has_result(self) -> bool:
        return self.result is not None

    def reset_results(self) -> None:
        self.result = None
        self.user_cluster_names.clear()
        self._dirty = False


__all__ = ["AnalysisSession"]
