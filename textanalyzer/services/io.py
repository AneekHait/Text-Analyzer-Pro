"""File I/O service — load, save, export helpers."""

from __future__ import annotations

import os
from typing import List, Optional

import joblib
import pandas as pd

from textanalyzer.engine.cluster import get_file_extension, get_sheet_names, load_table, save_results


class IOService:
    """Stateless façade over file-system operations."""

    @staticmethod
    def sheet_names(path: str) -> List[str]:
        return get_sheet_names(path)

    @staticmethod
    def load_table(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
        return load_table(path, sheet_name=sheet_name)

    @staticmethod
    def save_results(df: pd.DataFrame, path: str) -> str:
        return save_results(df, path)

    @staticmethod
    def save_model(model, vectorizer, cluster_names: dict, top_keywords: dict, path: str) -> None:
        joblib.dump(
            {
                "model": model,
                "vectorizer": vectorizer,
                "cluster_names": cluster_names,
                "top_keywords": top_keywords,
            },
            path,
        )

    @staticmethod
    def load_model(path: str) -> dict:
        """Load a joblib payload saved by save_model().

        Returns the raw payload dict so callers can pull out model / vectorizer /
        cluster_names / top_keywords / algorithm individually. Raises FileNotFoundError
        if the path doesn't exist, or RuntimeError if the payload is missing required
        keys.
        """
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Model file not found: {path}")
        payload = joblib.load(path)
        if not isinstance(payload, dict) or "model" not in payload or "vectorizer" not in payload:
            raise RuntimeError("Model payload missing required keys ('model', 'vectorizer').")
        return payload

    @staticmethod
    def default_output_path(input_path: str) -> str:
        base, ext = os.path.splitext(input_path)
        out_ext = ext if ext in {".csv", ".json"} else ".xlsx"
        return f"{base}_clustered{out_ext}"

    @staticmethod
    def file_extension(path: str) -> str:
        return get_file_extension(path)


__all__ = ["IOService"]
