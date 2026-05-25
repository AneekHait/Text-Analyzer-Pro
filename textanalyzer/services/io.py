"""File I/O service — load, save, export helpers."""

from __future__ import annotations

import os
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

import joblib
import pandas as pd

from textanalyzer.engine.cluster import (
    TaxonomyResult,
    get_file_extension,
    get_sheet_names,
    load_table,
    save_results,
)


TAXONOMY_SCHEMA_VERSION = 1
_CATEGORIZATION_COLUMNS = ("Repetitive/Non-Repetitive", "Subcategory", "Confidence")


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

    # ------------------------------------------------------------------
    # Taxonomy save/load (single-level subcategory bundles)
    # ------------------------------------------------------------------
    @staticmethod
    def save_taxonomy(
        result: TaxonomyResult,
        path: str,
        *,
        user_renames: Optional[Dict[str, str]] = None,
    ) -> None:
        """Persist a trained taxonomy as a .joblib bundle.

        The vectorizer carries its own ``__getstate__`` (EmbeddingVectorizer
        drops the heavy model handle on pickle) so the bundle stays small even
        when embedding mode was used. The result's ``manifest`` (model id,
        knobs, timestamps, scale path) is persisted so a future Load Taxonomy
        can surface "where did this come from" to the user.
        """
        joblib.dump(
            {
                "schema_version": TAXONOMY_SCHEMA_VERSION,
                "kind": "taxonomy",
                "vectorizer": result.vectorizer,
                "sub_centroids": result.sub_centroids,
                "subcategory_names": dict(result.subcategory_names),
                "sub_fingerprints": dict(result.sub_fingerprints),
                "avg_confidence_by_cluster": dict(result.avg_confidence_by_cluster),
                "samples_by_cluster": dict(result.samples_by_cluster),
                "user_renames": dict(user_renames or {}),
                "manifest": dict(result.manifest),
                "created_at": datetime.now(timezone.utc).isoformat(),
            },
            path,
        )

    @staticmethod
    def describe_taxonomy(payload: Dict[str, Any]) -> str:
        """Render a one-paragraph 'where did this taxonomy come from?' summary.

        Used by the Load Taxonomy flow to surface the manifest to the user
        before they apply a saved taxonomy to fresh data.
        """
        m = payload.get("manifest") or {}
        n_subs = len(payload.get("subcategory_names") or {})
        parts: List[str] = []
        if m.get("created_at"):
            parts.append(f"trained {m['created_at'][:10]}")
        if m.get("vectorizer_kind"):
            label = m["vectorizer_kind"]
            model = m.get("embedding_model")
            if model:
                label += f" ({model})"
            parts.append(label)
        if m.get("min_cluster_size") is not None:
            parts.append(f"min_cluster_size={m['min_cluster_size']}")
        if m.get("umap_applied"):
            parts.append("UMAP-reduced")
        if m.get("scale_path") and m["scale_path"] != "hdbscan":
            parts.append(m["scale_path"])
        if m.get("n_rows"):
            parts.append(f"{m['n_rows']} training rows")
        parts.append(f"{n_subs} subcategories")
        return "Taxonomy: " + ", ".join(parts) + "."

    @staticmethod
    def load_taxonomy(path: str) -> Dict[str, Any]:
        """Inverse of save_taxonomy.

        Raises FileNotFoundError if the path doesn't exist, RuntimeError if the
        payload is missing required keys or carries an unsupported schema
        version (so users see "taxonomy was saved with schema vN" rather than
        a silent mislabel).
        """
        if not os.path.isfile(path):
            raise FileNotFoundError(f"Taxonomy file not found: {path}")
        payload = joblib.load(path)
        if not isinstance(payload, dict):
            raise RuntimeError("Taxonomy payload is not a dict.")
        if payload.get("kind") != "taxonomy":
            raise RuntimeError(
                "File doesn't look like a taxonomy bundle (kind != 'taxonomy'). "
                "Use Load Model for clustering .joblib files."
            )
        version = int(payload.get("schema_version", 0) or 0)
        if version != TAXONOMY_SCHEMA_VERSION:
            raise RuntimeError(
                f"Taxonomy file was saved with schema v{version}, "
                f"this build expects v{TAXONOMY_SCHEMA_VERSION}."
            )
        for key in ("vectorizer", "sub_centroids", "subcategory_names"):
            if key not in payload:
                raise RuntimeError(f"Taxonomy payload missing required key: {key!r}")
        return payload

    # ------------------------------------------------------------------
    # Multi-sheet Excel export for categorization output
    # ------------------------------------------------------------------
    @staticmethod
    def _suggest_group_labels(
        centroids: Any, cluster_names: Dict[int, str], *, max_groups: int = 8,
    ) -> Dict[str, str]:
        """Suggest a coarse Group label for each subcategory via Ward linkage.

        Pure post-hoc grouping — runs hierarchical clustering on the centroid
        matrix, slices into ``max_groups`` clusters, and assigns each
        subcategory the most representative keyword of its group as the Group
        label. Best-effort: returns an empty dict on any failure so the pivot
        gracefully degrades to the no-grouping path.
        """
        try:
            import numpy as _np
            from scipy.cluster.hierarchy import fcluster, linkage  # type: ignore[import-untyped]
        except Exception:
            return {}
        if centroids is None or len(centroids) < 2:
            return {}
        try:
            arr = _np.asarray(centroids, dtype=_np.float64)
            if arr.ndim != 2 or arr.shape[0] < 2:
                return {}
            n = arr.shape[0]
            n_groups = max(2, min(max_groups, n))
            # Cosine-distance via 1 - cosine_similarity. Ward needs Euclidean,
            # so we use the unit-normalized vectors and Ward on those — gives
            # a sensible hierarchical structure on embedding centroids.
            norms = _np.linalg.norm(arr, axis=1)
            norms[norms == 0] = 1.0
            unit = arr / norms[:, None]
            Z = linkage(unit, method="ward")
            flat = fcluster(Z, t=n_groups, criterion="maxclust")
        except Exception:
            return {}
        # Map cluster_id (in centroid order) → flat group → representative name.
        ordered_cids = sorted(cluster_names.keys())
        if len(ordered_cids) != n:
            return {}
        group_to_members: Dict[int, List[int]] = {}
        for cid, g in zip(ordered_cids, flat.tolist()):
            group_to_members.setdefault(int(g), []).append(int(cid))
        # Group label = first 1-2 words of the largest-cluster name in the group.
        out: Dict[str, str] = {}
        for members in group_to_members.values():
            largest = members[0]  # arbitrary tiebreak; deterministic by sort above
            label_seed = cluster_names.get(largest, f"Group {largest}")
            label = " ".join(label_seed.split()[:2]) or label_seed
            for cid in members:
                out[cluster_names.get(cid, f"Subcluster {cid}")] = label
        return out

    @staticmethod
    def build_pivot_dataframe(
        df: pd.DataFrame,
        *,
        avg_confidence_by_subcat: Optional[Dict[str, float]] = None,
        group_labels: Optional[Dict[str, str]] = None,
    ) -> pd.DataFrame:
        """Build a flat rollup DataFrame mirroring the reference workbook's pivot.

        Base schema: ``Row Labels | Count | %``. Adds an ``Avg Confidence``
        column when ``avg_confidence_by_subcat`` is supplied, and a leading
        ``Group`` column when ``group_labels`` is supplied (rows sort by Group
        then descending count within group).
        """
        if not all(col in df.columns for col in _CATEGORIZATION_COLUMNS):
            raise ValueError(
                "build_pivot_dataframe requires the three categorization columns: "
                f"{_CATEGORIZATION_COLUMNS}"
            )
        total = len(df)
        cols = ["Row Labels", "Count", "%"]
        if avg_confidence_by_subcat:
            cols.append("Avg Confidence")
        if group_labels:
            cols.insert(0, "Group")
        if total == 0:
            return pd.DataFrame(columns=cols)

        rep_mask = df["Repetitive/Non-Repetitive"] == "Repetitive"
        rep_count = int(rep_mask.sum())
        non_rep_count = total - rep_count

        sub_counts = (
            df.loc[rep_mask, "Subcategory"]
            .value_counts()
            .sort_values(ascending=False)
        )

        def _row(label: str, count: int, pct: float, *, subcat: Optional[str] = None,
                 group: Optional[str] = None) -> Dict[str, Any]:
            row: Dict[str, Any] = {"Row Labels": label, "Count": count, "%": pct}
            if avg_confidence_by_subcat is not None:
                row["Avg Confidence"] = (
                    float(avg_confidence_by_subcat.get(subcat, 0.0)) if subcat else ""
                )
            if group_labels is not None:
                row["Group"] = group if group is not None else ""
            return row

        rows: List[Dict[str, Any]] = []
        rows.append(_row("Repetitive", rep_count, rep_count / total))
        # Order by (group, -count) when grouping is requested; pure -count otherwise.
        if group_labels:
            items = sorted(
                sub_counts.items(),
                key=lambda kv: (group_labels.get(str(kv[0]), "zzz"), -int(kv[1])),
            )
        else:
            items = list(sub_counts.items())
        for name, count in items:
            count_i = int(count)
            pct = count_i / total
            group = group_labels.get(str(name)) if group_labels else None
            rows.append(_row(f"  {name}", count_i, pct, subcat=str(name), group=group))
        rows.append(_row("Non-Repetitive", non_rep_count, non_rep_count / total))
        rows.append(_row("TOTAL", total, 1.0))
        return pd.DataFrame(rows, columns=cols)

    @staticmethod
    def save_results_with_pivot(
        df: pd.DataFrame, out_path: str, *,
        sheet_name: str = "Inc",
        taxonomy_result: Optional[TaxonomyResult] = None,
        include_groups: bool = True,
    ) -> str:
        """Write the dataframe + an auto-generated `pivot` sheet to .xlsx.

        When ``taxonomy_result`` is supplied, the pivot gains an
        ``Avg Confidence`` column (from ``taxonomy_result.avg_confidence_by_cluster``)
        and — when ``include_groups`` is true and SciPy is available — a
        leading ``Group`` column derived from Ward hierarchical clustering on
        the cluster centroids.

        For non-xlsx outputs (.csv / .json), falls back to plain save_results —
        those formats don't carry a second sheet.
        """
        ext = get_file_extension(out_path)
        if ext != ".xlsx":
            return save_results(df, out_path)

        avg_conf_by_subcat: Optional[Dict[str, float]] = None
        group_labels: Optional[Dict[str, str]] = None
        if taxonomy_result is not None:
            avg_conf_by_subcat = {
                taxonomy_result.subcategory_names.get(int(cid), str(cid)): float(v)
                for cid, v in (taxonomy_result.avg_confidence_by_cluster or {}).items()
            }
            if include_groups:
                group_labels = IOService._suggest_group_labels(
                    taxonomy_result.sub_centroids,
                    taxonomy_result.subcategory_names,
                ) or None
        pivot_df = IOService.build_pivot_dataframe(
            df,
            avg_confidence_by_subcat=avg_conf_by_subcat,
            group_labels=group_labels,
        )

        def _write(target_path: str) -> None:
            with pd.ExcelWriter(target_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                pivot_df.to_excel(writer, index=False, sheet_name="pivot")

        try:
            _write(out_path)
            print(f"Saved results + pivot to {out_path}")
            return out_path
        except PermissionError:
            base, original_ext = os.path.splitext(out_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            alt_path = f"{base}_writable_{timestamp}{original_ext}"
            _write(alt_path)
            print(f"Could not overwrite {out_path} (permission denied). Saved to {alt_path}.")
            return alt_path


__all__ = ["IOService", "TAXONOMY_SCHEMA_VERSION"]
