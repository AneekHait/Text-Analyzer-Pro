#!/usr/bin/env python3
"""
cluster_tool.py

A CLI tool to cluster text data from Excel, CSV, or JSON files and write cluster labels back to a table file.

Features:
- Load a table file and allow user to choose a text column
- Preprocessing: lowercasing, basic normalization (non-string -> string), drop/handle missing
- TF-IDF vectorization (with sklearn, english stop words)
- Clustering: KMeans, DBSCAN, Agglomerative
- Optional visualization (PCA or t-SNE)
- Top keywords per cluster summary
- Save trained model (joblib)

Usage examples:
  python cluster_tool.py --input data.xlsx --column comments --algorithm kmeans --n_clusters 5 --output clustered.xlsx
  python cluster_tool.py --input data.csv --column comments --algorithm kmeans --n_clusters 5 --output clustered.csv
"""

import argparse
import hashlib
import os
import re
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import (
    ENGLISH_STOP_WORDS,
    HashingVectorizer,
    TfidfTransformer,
    TfidfVectorizer,
)
from sklearn.cluster import (
    AgglomerativeClustering,
    DBSCAN,
    KMeans,
    MiniBatchKMeans,
)
from sklearn.decomposition import PCA
from sklearn.manifold import TSNE
from sklearn.metrics import (
    calinski_harabasz_score,
    davies_bouldin_score,
    silhouette_score,
)
from sklearn.neighbors import NearestNeighbors
from scipy.sparse import vstack as sparse_vstack
import joblib

# matplotlib + seaborn are heavy (~300-500 ms cold start on Windows) and only
# needed by visualize_embeddings(); imported lazily inside that function.
# Same for nltk — only the optional lemmatize path touches it.

# Optional GPU-accelerated KMeans (cuML). The module remains importable without it.
try:
    import cuml  # type: ignore[import-untyped]
    _CUML_AVAILABLE = True
except Exception:
    _CUML_AVAILABLE = False

# Optional HDBSCAN — automatic cluster-count selection.
try:
    import hdbscan  # type: ignore[import-untyped]
    _HDBSCAN_AVAILABLE = True
except ImportError:
    _HDBSCAN_AVAILABLE = False

# Optional sentence-transformers — dense semantic embeddings as an alternative
# to TF-IDF. Heavy (pulls torch); kept optional. Import is checked at top-level
# so the GUI can disable the embedding option early, but the actual
# ``SentenceTransformer`` import is deferred to ``EmbeddingVectorizer`` so
# cold-start of this module stays light.
try:
    import sentence_transformers  # type: ignore[import-untyped]  # noqa: F401
    _ST_AVAILABLE = True
except ImportError:
    _ST_AVAILABLE = False

# Optional UMAP — dim-reduction step that goes between sentence embeddings and
# HDBSCAN. The BERTopic / Top2Vec recipe: reducing 384-dim embeddings to ~15
# dims kills distance concentration in high-D space and produces noticeably
# finer-grained, less-noisy clusters. Pure speedup + quality win when present.
try:
    import umap  # type: ignore[import-untyped]  # noqa: F401
    _UMAP_AVAILABLE = True
except Exception:  # umap-learn may fail to import on partial installs
    _UMAP_AVAILABLE = False

DEFAULT_EMBEDDING_MODEL = "sentence-transformers/all-MiniLM-L6-v2"

# At >= this many rows we switch sub-clustering from HDBSCAN to MiniBatchKMeans
# with percentile-based noise tagging — HDBSCAN's worst-case runtime starts
# dominating the workflow above this scale.
_CATEGORIZATION_SCALE_GUARD = 25_000

# Cap silhouette computations on huge datasets so metrics stay tractable.
_MAX_SILHOUETTE_SAMPLES = 5000

# Cleaning pipeline lives in its own module; re-exported here for compat.
from textanalyzer.engine.cleaning import (  # noqa: F401
    TextCleaningConfig,
    TextCleaningResult,
    clean_text_value,
    coerce_text_column,
    get_default_text_cleaning_config,
    prepare_text_cleaning,
    preprocess_texts,
)


def _to_dense(X):
    """Densify a sparse matrix; pass through if X is already dense."""
    return X.toarray() if hasattr(X, "toarray") else X


# Excel-family inputs handled via pd.read_excel. Pandas auto-picks the right
# engine based on extension: openpyxl for .xlsx/.xlsm/.xltx/.xltm, xlrd for
# .xls, pyxlsb for .xlsb, odf (odfpy) for .ods. The corresponding packages
# must be installed — see requirements.txt.
EXCEL_INPUT_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".ods"}
SUPPORTED_INPUT_EXTENSIONS = {*EXCEL_INPUT_EXTENSIONS, ".csv", ".json"}
SINGLE_TABLE_SHEET_NAME = "Data"


def get_file_extension(path: str) -> str:
    return os.path.splitext(path)[1].lower()


def _coerce_json_frame(raw_json: Any) -> pd.DataFrame:
    if isinstance(raw_json, pd.DataFrame):
        return raw_json
    if isinstance(raw_json, pd.Series):
        return raw_json.to_frame()
    if isinstance(raw_json, list):
        return pd.json_normalize(raw_json)
    if isinstance(raw_json, dict):
        return pd.json_normalize(raw_json)
    raise ValueError("Unsupported JSON structure. Expected an object, array, or tabular JSON payload.")


_TEXT_ENCODING_FALLBACKS = ("utf-8", "utf-8-sig", "cp1252", "latin-1")


def _read_with_encoding_fallback(reader, path: str, **kwargs):
    """Try `reader(path, encoding=...)` across the standard fallback chain.

    Excel exports on Windows are typically cp1252 — bytes like 0x96 (en-dash)
    blow up under utf-8. We try utf-8 first (fast happy path) then utf-8-sig
    (handles BOM), then cp1252 (Western Windows default), then latin-1 (every
    byte is valid latin-1, so this never raises). Warns when a non-utf-8
    encoding wins so the user knows their file isn't UTF-8.
    """
    last_exc: Optional[UnicodeDecodeError] = None
    for encoding in _TEXT_ENCODING_FALLBACKS:
        try:
            df = reader(path, encoding=encoding, **kwargs)
        except UnicodeDecodeError as exc:
            last_exc = exc
            continue
        if encoding != "utf-8":
            warnings.warn(
                f"Loaded {os.path.basename(path)} as {encoding} "
                f"(file is not UTF-8). Re-save as UTF-8 to silence this."
            )
        return df
    # latin-1 should make the loop above unreachable, but re-raise just in case.
    if last_exc is not None:
        raise last_exc
    raise RuntimeError(f"Could not decode {path} with any known encoding")


_UNSUPPORTED_MSG = (
    "Unsupported file type '{ext}'. Supported types: Excel "
    "(.xlsx, .xlsm, .xltx, .xltm, .xls, .xlsb, .ods), CSV (.csv), JSON (.json)."
)


def load_table(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    ext = get_file_extension(path)
    if ext in EXCEL_INPUT_EXTENSIONS:
        # Engine=None lets pandas auto-pick the right engine for the format.
        # openpyxl → .xlsx/.xlsm/.xltx/.xltm, xlrd → .xls, pyxlsb → .xlsb,
        # odf → .ods. Each engine must be installed (see requirements.txt).
        if sheet_name is None:
            return pd.read_excel(path)
        return pd.read_excel(path, sheet_name=sheet_name)
    if ext == ".csv":
        return _read_with_encoding_fallback(pd.read_csv, path)
    if ext == ".json":
        try:
            return _coerce_json_frame(_read_with_encoding_fallback(pd.read_json, path))
        except ValueError:
            return _coerce_json_frame(
                _read_with_encoding_fallback(pd.read_json, path, lines=True)
            )
    raise ValueError(_UNSUPPORTED_MSG.format(ext=ext or "(none)"))


def get_sheet_names(path: str) -> List[str]:
    ext = get_file_extension(path)
    if ext in EXCEL_INPUT_EXTENSIONS:
        workbook = pd.ExcelFile(path)
        return workbook.sheet_names
    if ext in {".csv", ".json"}:
        return [SINGLE_TABLE_SHEET_NAME]
    raise ValueError(_UNSUPPORTED_MSG.format(ext=ext or "(none)"))


def load_excel(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    return load_table(path, sheet_name=sheet_name)




class EmbeddingVectorizer:
    """Dense sentence-embedding vectorizer with the same shape as TfidfVectorizer.

    Wraps a ``sentence_transformers.SentenceTransformer`` so the rest of the
    pipeline can keep treating the vectorizer as an opaque ``(.fit_transform,
    .transform)`` pair. Embeddings are L2-normalized so KMeans on Euclidean
    distance behaves like spherical KMeans (which is the standard for dense
    embedding clustering) and DBSCAN's existing ``metric="cosine"`` stays
    coherent.

    A side ``TfidfVectorizer`` is fit on the *same* training texts purely so
    ``get_top_keywords_per_cluster`` can still produce readable cluster names.

    Persistence: ``__getstate__`` drops the heavy ``_model`` handle so saved
    ``.joblib`` files stay a few KB. The model is re-instantiated lazily on
    the next ``transform()`` call and pulled from the sentence-transformers
    on-disk cache.
    """

    def __init__(
        self,
        model_name: str = DEFAULT_EMBEDDING_MODEL,
        device: str = "cpu",
        batch_size: int = 32,
    ):
        self.model_name = model_name
        self.device = device
        self.batch_size = batch_size
        self._model = None
        self._side_tfidf: Optional[TfidfVectorizer] = None
        self._training_texts: Optional[List[str]] = None

    def _ensure_loaded(self):
        if self._model is None:
            if not _ST_AVAILABLE:
                raise ImportError(
                    "sentence-transformers not installed. "
                    "Install with: pip install sentence-transformers"
                )
            from sentence_transformers import SentenceTransformer  # type: ignore[import-untyped]
            self._model = SentenceTransformer(self.model_name, device=self.device)

    def _encode(self, texts: List[str]) -> np.ndarray:
        self._ensure_loaded()
        if not texts:
            return np.zeros((0, 0), dtype=np.float32)
        return np.asarray(
            self._model.encode(
                list(texts),
                batch_size=self.batch_size,
                show_progress_bar=False,
                convert_to_numpy=True,
                normalize_embeddings=True,
            )
        )

    def fit_transform(self, texts: List[str]) -> np.ndarray:
        texts = list(texts)
        X = self._encode(texts)
        # Side TF-IDF for cluster-naming. Failure here is non-fatal: keyword
        # extraction will fall back to empty lists, like the hashing path.
        try:
            self._side_tfidf = TfidfVectorizer(stop_words="english", max_features=2000)
            self._side_tfidf.fit(texts)
            self._training_texts = texts
        except Exception:
            self._side_tfidf = None
            self._training_texts = None
        return X

    def transform(self, texts: List[str]) -> np.ndarray:
        return self._encode(list(texts))

    def get_feature_names_out(self):
        if self._side_tfidf is not None:
            return self._side_tfidf.get_feature_names_out()
        raise AttributeError("EmbeddingVectorizer has no feature names without a side TF-IDF")

    def __getstate__(self):
        state = self.__dict__.copy()
        state["_model"] = None
        return state


def vectorize_texts(
    texts: List[str],
    max_features: Optional[int] = 2000,
    *,
    use_hashing: bool = False,
    chunk_size: int = 10000,
    min_df: int = 1,
    max_df: float = 1.0,
    ngram_range: Tuple[int, int] = (1, 1),
    custom_stopwords: Optional[List[str]] = None,
    vectorizer_kind: str = "tfidf",
    embedding_model: str = DEFAULT_EMBEDDING_MODEL,
    embedding_device: str = "cpu",
    embedding_batch_size: int = 32,
) -> Tuple[Any, Any]:
    """Vectorize texts into a feature matrix.

    Default behavior (``vectorizer_kind="tfidf"`` + ``use_hashing=False``) is
    unchanged from the previous signature.

    Modes:
      - ``vectorizer_kind="tfidf"``: TF-IDF (default). Sparse output. Use
        ``use_hashing=True`` for a stateless, memory-bounded path that streams
        documents in chunks; the returned vectorizer is a ``HashingVectorizer``
        and ``get_top_keywords_per_cluster`` returns empty keyword lists since
        hashing has no feature names.
      - ``vectorizer_kind="embedding"``: dense sentence embeddings via
        sentence-transformers. Requires ``pip install sentence-transformers``;
        raises ``ImportError`` with an install hint if missing. Returns an
        ``EmbeddingVectorizer`` and an L2-normalized dense ``np.ndarray``.
    """
    if vectorizer_kind == "embedding":
        if use_hashing:
            raise ValueError(
                "use_hashing is incompatible with vectorizer_kind='embedding'. "
                "Pick one vectorization mode."
            )
        vec = EmbeddingVectorizer(
            model_name=embedding_model,
            device=embedding_device,
            batch_size=embedding_batch_size,
        )
        return vec, vec.fit_transform(texts)

    if vectorizer_kind != "tfidf":
        raise ValueError(
            f"Unknown vectorizer_kind={vectorizer_kind!r}. Choose 'tfidf' or 'embedding'."
        )

    stop_words: Any = "english"
    if custom_stopwords:
        stop_words = list(ENGLISH_STOP_WORDS) + list(custom_stopwords)

    if not use_hashing:
        vectorizer = TfidfVectorizer(
            stop_words=stop_words,
            max_features=max_features,
            min_df=min_df,
            max_df=max_df,
            ngram_range=ngram_range,
        )
        X = vectorizer.fit_transform(texts)
        return vectorizer, X

    hashing = HashingVectorizer(
        n_features=max_features or 2 ** 20,
        alternate_sign=False,
        ngram_range=ngram_range,
    )
    transformer = TfidfTransformer()
    chunks = []
    for start in range(0, len(texts), chunk_size):
        chunks.append(hashing.transform(texts[start : start + chunk_size]))
    X_counts = sparse_vstack(chunks) if chunks else hashing.transform([])
    X = transformer.fit_transform(X_counts)
    return hashing, X


def cluster_texts(
    X,
    algorithm: str = "kmeans",
    n_clusters: int = 5,
    eps: float = 0.5,
    min_samples: int = 5,
    random_state: int = 42,
    *,
    min_cluster_size: int = 5,
    use_gpu: bool = False,
    minibatch_threshold: int = 10000,
):
    """Cluster a feature matrix.

    Algorithms: 'kmeans', 'dbscan', 'agglomerative', 'hdbscan'.

    For KMeans, automatically switches to MiniBatchKMeans when the dataset has
    `>= minibatch_threshold` rows (default 10k). When `use_gpu=True` and cuML
    is installed, runs KMeans on GPU; otherwise falls back to CPU silently.
    """
    n_samples = X.shape[0]
    if algorithm in ("kmeans", "agglomerative"):
        if n_clusters > n_samples:
            raise ValueError(
                f"The number of clusters ({n_clusters}) cannot be greater than the number of samples ({n_samples})."
            )

    if algorithm == "agglomerative" and n_samples > 15000:
        raise ValueError(
            f"Agglomerative clustering is not recommended for datasets with more than 15,000 samples "
            f"(current: {n_samples}). Consider KMeans or DBSCAN."
        )

    if algorithm == "kmeans":
        if use_gpu and _CUML_AVAILABLE:
            try:
                X_dense = _to_dense(X).astype("float32")
                # cuML's KMeans defaults to scalable-k-means++ (a parallel variant)
                # so initialization quality matches the CPU path.
                model = cuml.cluster.KMeans(n_clusters=n_clusters, random_state=random_state)
                model.fit(X_dense)
                return model, model.predict(X_dense)
            except Exception:
                pass  # fall through to CPU path
        # Both branches use kmeans++ (sklearn's default) — set explicitly so the
        # init algorithm is visible at the call site.
        if n_samples >= minibatch_threshold:
            model = MiniBatchKMeans(
                n_clusters=n_clusters, init="k-means++", n_init=10, random_state=random_state
            )
        else:
            model = KMeans(
                n_clusters=n_clusters, init="k-means++", n_init=10, random_state=random_state
            )
        labels = model.fit_predict(X)
        return model, labels

    if algorithm == "dbscan":
        model = DBSCAN(eps=eps, min_samples=min_samples, metric="cosine")
        return model, model.fit_predict(X)

    if algorithm == "hdbscan":
        if not _HDBSCAN_AVAILABLE:
            raise ImportError("HDBSCAN not installed. Install with: pip install hdbscan")
        X_dense = _to_dense(X)
        model = hdbscan.HDBSCAN(
            min_cluster_size=min_cluster_size,
            min_samples=min_samples,
            metric="euclidean",
            cluster_selection_method="eom",
        )
        return model, model.fit_predict(X_dense)

    if algorithm == "agglomerative":
        model = AgglomerativeClustering(n_clusters=n_clusters)
        return model, model.fit_predict(_to_dense(X))

    raise ValueError(f"Unknown algorithm: {algorithm}")


# ============================================================================
# Validation, k-selection, algorithm comparison, predict-wrapper
# ============================================================================

@dataclass
class ValidationResult:
    """Outcome of validate_input()."""

    is_valid: bool
    errors: List[str]
    warnings: List[str]
    stats: Dict[str, Any]


def validate_input(
    df: pd.DataFrame,
    column: str,
    algorithm: str = "kmeans",
    n_clusters: int = 5,
) -> ValidationResult:
    """Validate a dataframe + column choice before clustering.

    Returns errors, warnings, and a stats dict with row counts, null/empty
    percentages, unique-text count, and text-length distribution. Designed for
    a "Are you sure?" panel that shows users what they're about to cluster.
    """
    errors: List[str] = []
    warnings_list: List[str] = []
    stats: Dict[str, Any] = {}

    if column not in df.columns:
        errors.append(f"Column '{column}' not found. Available: {list(df.columns)}")
        return ValidationResult(False, errors, warnings_list, stats)

    col_data = df[column]
    total_rows = len(col_data)
    stats["total_rows"] = total_rows

    null_count = int(col_data.isna().sum())
    stats["null_count"] = null_count
    stats["null_percentage"] = round(null_count / total_rows * 100, 1) if total_rows > 0 else 0

    text_data = col_data.fillna("").astype(str)
    empty_count = int((text_data.str.strip() == "").sum())
    stats["empty_count"] = empty_count
    stats["empty_percentage"] = round(empty_count / total_rows * 100, 1) if total_rows > 0 else 0

    non_empty = total_rows - empty_count
    stats["non_empty_count"] = non_empty
    if non_empty == 0:
        errors.append("All values in the selected column are empty or null")
        return ValidationResult(False, errors, warnings_list, stats)

    if empty_count > total_rows * 0.5:
        warnings_list.append(f"More than 50% of values are empty ({stats['empty_percentage']}%)")

    unique_texts = int(text_data[text_data.str.strip() != ""].nunique())
    stats["unique_texts"] = unique_texts

    if algorithm in ("kmeans", "agglomerative"):
        if n_clusters > non_empty:
            errors.append(
                f"n_clusters ({n_clusters}) cannot be greater than non-empty samples ({non_empty})"
            )
        elif n_clusters > unique_texts:
            warnings_list.append(
                f"n_clusters ({n_clusters}) > unique texts ({unique_texts}). "
                "Some clusters may be empty or identical."
            )

    text_lengths = text_data[text_data.str.strip() != ""].str.len()
    stats["avg_text_length"] = round(float(text_lengths.mean()), 1) if len(text_lengths) > 0 else 0
    stats["min_text_length"] = int(text_lengths.min()) if len(text_lengths) > 0 else 0
    stats["max_text_length"] = int(text_lengths.max()) if len(text_lengths) > 0 else 0

    if stats["avg_text_length"] and stats["avg_text_length"] < 10:
        warnings_list.append(
            f"Average text length is very short ({stats['avg_text_length']} chars). "
            "Clustering may not be meaningful."
        )

    if algorithm == "agglomerative" and non_empty > 10000:
        warnings_list.append(
            f"Agglomerative clustering with {non_empty} samples may use significant memory. "
            "Consider KMeans for large datasets."
        )
    if algorithm == "dbscan":
        warnings_list.append(
            "DBSCAN may produce many noise points (label -1) depending on eps/min_samples."
        )

    return ValidationResult(len(errors) == 0, errors, warnings_list, stats)


def find_optimal_k(
    X,
    k_range: Tuple[int, int] = (2, 15),
    method: str = "silhouette",
    random_state: int = 42,
) -> Dict[str, Any]:
    """Search a k-range for the most natural cluster count.

    method='silhouette' picks the k whose KMeans labelling has the highest
    silhouette score (with confidence band based on the absolute score).
    method='elbow' picks the k where inertia bends most sharply.

    Returns a dict with `optimal_k`, `scores` (per-k silhouette), `inertias`,
    `recommendation` (human-readable string), `confidence`.
    """
    k_min, k_max = k_range
    n_samples = X.shape[0]
    k_max = min(k_max, n_samples - 1)
    if k_max < k_min:
        return {
            "optimal_k": 2,
            "scores": {},
            "inertias": {},
            "recommendation": f"Not enough samples ({n_samples}) for cluster analysis",
            "confidence": "low",
        }

    scores: Dict[int, float] = {}
    inertias: Dict[int, float] = {}
    previous_score: Optional[float] = None
    declining = 0

    for k in range(k_min, k_max + 1):
        try:
            model = KMeans(n_clusters=k, init="k-means++", n_init=10, random_state=random_state)
            labels = model.fit_predict(X)
            if len(set(labels)) < 2:
                continue
            if X.shape[0] > _MAX_SILHOUETTE_SAMPLES:
                idx = np.random.RandomState(random_state).choice(
                    X.shape[0], _MAX_SILHOUETTE_SAMPLES, replace=False
                )
                score = silhouette_score(X[idx], labels[idx])
            else:
                score = silhouette_score(X, labels)
            scores[k] = float(score)
            inertias[k] = float(model.inertia_)
            if previous_score is not None and score < previous_score:
                declining += 1
                if declining >= 3 and len(scores) >= 5:
                    break
            else:
                declining = 0
            previous_score = score
        except Exception:
            continue

    if not scores:
        return {
            "optimal_k": 2,
            "scores": {},
            "inertias": {},
            "recommendation": "Could not compute cluster scores",
            "confidence": "low",
        }

    if method == "silhouette":
        optimal_k = max(scores, key=scores.get)
        best = scores[optimal_k]
        if best >= 0.5:
            confidence, quality = "high", "strong"
        elif best >= 0.3:
            confidence, quality = "medium", "moderate"
        elif best >= 0.1:
            confidence, quality = "low", "weak"
        else:
            confidence, quality = "very_low", "no clear"
        recommendation = (
            f"Recommended: {optimal_k} clusters (silhouette {best:.3f}). "
            f"This indicates {quality} cluster structure."
        )
    else:  # elbow
        ks = sorted(inertias.keys())
        if len(ks) < 3:
            optimal_k = ks[0] if ks else 2
        else:
            values = np.array([inertias[k] for k in ks])
            second = np.diff(np.diff(values))
            optimal_k = ks[min(int(np.argmax(second)) + 1, len(ks) - 1)]
        confidence = "medium"
        recommendation = f"Recommended: {optimal_k} clusters (elbow method)"

    return {
        "optimal_k": optimal_k,
        "scores": scores,
        "inertias": inertias,
        "recommendation": recommendation,
        "confidence": confidence,
    }


@dataclass
class AlgorithmResult:
    """Per-algorithm outcome from compare_algorithms()."""

    name: str
    labels: np.ndarray
    n_clusters: int
    silhouette: Optional[float]
    calinski_harabasz: Optional[float]
    davies_bouldin: Optional[float]
    noise_count: int
    runtime_seconds: float


def compare_algorithms(
    X,
    n_clusters: int = 5,
    eps: float = 0.5,
    min_samples: int = 5,
    random_state: int = 42,
    progress_callback: Optional[Callable[[str, float], None]] = None,
) -> List[AlgorithmResult]:
    """Run KMeans / DBSCAN / Agglomerative on the same matrix and report metrics.

    Each result includes silhouette, Calinski-Harabasz, Davies-Bouldin, and
    runtime — pick whichever you trust. Skips agglomerative on large data
    (>10k rows) since it would OOM.
    """
    import time

    results: List[AlgorithmResult] = []
    algorithms = ["kmeans", "dbscan", "agglomerative"]

    for i, alg in enumerate(algorithms):
        if progress_callback:
            progress_callback(f"Running {alg}...", i / len(algorithms))
        start = time.time()
        try:
            if alg == "kmeans":
                model = KMeans(
                    n_clusters=n_clusters, init="k-means++", n_init=10, random_state=random_state
                )
                labels = model.fit_predict(X)
            elif alg == "dbscan":
                model = DBSCAN(eps=eps, min_samples=min_samples, metric="cosine")
                labels = model.fit_predict(X)
            else:  # agglomerative
                if X.shape[0] > 10000:
                    results.append(AlgorithmResult(
                        name=alg,
                        labels=np.full(X.shape[0], -1),
                        n_clusters=0,
                        silhouette=None,
                        calinski_harabasz=None,
                        davies_bouldin=None,
                        noise_count=X.shape[0],
                        runtime_seconds=0.0,
                    ))
                    continue
                model = AgglomerativeClustering(n_clusters=n_clusters)
                labels = model.fit_predict(_to_dense(X))

            runtime = time.time() - start
            unique = set(labels)
            unique.discard(-1)
            actual_k = len(unique)
            noise = int((labels == -1).sum())

            silhouette = calinski = davies = None
            if actual_k >= 2:
                mask = labels != -1
                if mask.sum() >= 2:
                    Xm = X[mask]
                    labelsm = labels[mask]
                    Xm_dense = _to_dense(Xm)
                    try:
                        if Xm_dense.shape[0] > _MAX_SILHOUETTE_SAMPLES:
                            idx = np.random.RandomState(random_state).choice(
                                Xm_dense.shape[0], _MAX_SILHOUETTE_SAMPLES, replace=False
                            )
                            silhouette = float(silhouette_score(Xm_dense[idx], labelsm[idx]))
                        else:
                            silhouette = float(silhouette_score(Xm_dense, labelsm))
                    except Exception:
                        pass
                    try:
                        calinski = float(calinski_harabasz_score(Xm_dense, labelsm))
                    except Exception:
                        pass
                    try:
                        davies = float(davies_bouldin_score(Xm_dense, labelsm))
                    except Exception:
                        pass

            results.append(AlgorithmResult(
                name=alg,
                labels=labels,
                n_clusters=actual_k,
                silhouette=silhouette,
                calinski_harabasz=calinski,
                davies_bouldin=davies,
                noise_count=noise,
                runtime_seconds=runtime,
            ))
        except Exception:
            results.append(AlgorithmResult(
                name=alg,
                labels=np.array([]),
                n_clusters=0,
                silhouette=None,
                calinski_harabasz=None,
                davies_bouldin=None,
                noise_count=0,
                runtime_seconds=0.0,
            ))

    if progress_callback:
        progress_callback("Complete", 1.0)
    return results


def get_best_algorithm(results: List[AlgorithmResult]) -> Optional[str]:
    """Pick the highest-silhouette algorithm from a compare_algorithms() result list."""
    valid = [r for r in results if r.silhouette is not None]
    if not valid:
        return None
    return max(valid, key=lambda r: r.silhouette).name


class ApplicableModel:
    """Wrap a fitted clustering model so it has a working .predict() on new data.

    KMeans / MiniBatchKMeans delegate to the model's own .predict(). DBSCAN /
    Agglomerative / HDBSCAN don't natively support out-of-sample prediction,
    so the wrapper builds a NearestNeighbors index over the training matrix
    and labels new samples by their closest training neighbor's cluster.
    """

    _NN_METRIC = "cosine"

    def __init__(self, model, X_train, labels: np.ndarray, algorithm: str):
        self.model = model
        self.algorithm = algorithm
        self.labels = np.asarray(labels)
        self._nn = None
        self._train_labels: Optional[np.ndarray] = None

        if algorithm in ("dbscan", "agglomerative", "hdbscan"):
            self._nn = NearestNeighbors(n_neighbors=1, metric=self._NN_METRIC)
            self._nn.fit(_to_dense(X_train))
            self._train_labels = self.labels

    def predict(self, X):
        if self.algorithm in ("kmeans",) and hasattr(self.model, "predict"):
            return self.model.predict(X)
        if self._nn is not None and self._train_labels is not None:
            _, indices = self._nn.kneighbors(_to_dense(X))
            return self._train_labels[indices.flatten()]
        raise RuntimeError(f"Cannot predict with algorithm: {self.algorithm}")

    def fit_predict(self, X):
        return self.labels


def wrap_model_for_prediction(model, X_train, labels, algorithm: str) -> ApplicableModel:
    """Convenience wrapper around `ApplicableModel(...)`."""
    return ApplicableModel(model, X_train, labels, algorithm)


def get_top_keywords_per_cluster(
    vectorizer, X, labels: np.ndarray, top_n: int = 10
) -> Dict[int, List[Tuple[str, float]]]:
    """Compute top-N TF-IDF keywords per cluster.

    Works with TfidfVectorizer. If the vectorizer is HashingVectorizer (no
    feature names), returns empty keyword lists per cluster and warns once.

    For ``EmbeddingVectorizer``, swaps in the bundled side TF-IDF fit on the
    training texts so cluster names remain human-readable even when the
    primary feature matrix is dense embeddings.
    """
    result: Dict[int, List[Tuple[str, float]]] = {}

    # Embedding mode: keywords come from the side TF-IDF over the same texts.
    if isinstance(vectorizer, EmbeddingVectorizer):
        if vectorizer._side_tfidf is None or vectorizer._training_texts is None:
            warnings.warn(
                "EmbeddingVectorizer has no side TF-IDF — top keywords unavailable."
            )
            for label in sorted(np.unique(labels)):
                if label == -1:
                    continue
                result[int(label)] = []
            return result
        side_X = vectorizer._side_tfidf.transform(vectorizer._training_texts)
        return get_top_keywords_per_cluster(
            vectorizer._side_tfidf, side_X, labels, top_n=top_n
        )

    try:
        features = vectorizer.get_feature_names_out()
    except Exception:
        warnings.warn(
            "Vectorizer does not expose feature names (e.g. HashingVectorizer). "
            "Top keywords will be unavailable."
        )
        for label in sorted(np.unique(labels)):
            if label == -1:
                continue
            result[int(label)] = []
        return result

    for label in sorted(np.unique(labels)):
        if label == -1:
            continue
        mask = labels == label
        if not np.any(mask):
            result[int(label)] = []
            continue
        cluster_X = X[mask]
        mean_tfidf = np.asarray(cluster_X.mean(axis=0)).flatten()
        top_idx = mean_tfidf.argsort()[-top_n:][::-1]
        result[int(label)] = list(zip(features[top_idx].tolist(), mean_tfidf[top_idx].tolist()))
    return result


def assign_cluster_names(top_keywords: Dict[int, List[Tuple[str, float]]], name_top_n: int = 3, joiner: str = ", ") -> Dict[int, str]:
    """Create a simple, descriptive name for each cluster from its top keywords.

    - top_keywords: mapping cluster_id -> list of (term, score)
    - name_top_n: how many top terms to include in the name
    - joiner: string used to join terms in the cluster name
    Returns mapping cluster_id -> cluster_name
    """
    names = {}
    for cid, terms in top_keywords.items():
        if not terms:
            names[cid] = f"cluster_{cid}"
            continue
        top_terms = [t for t, s in terms][:name_top_n]
        # sanitize and join
        safe_terms = [str(t).replace(" ", "_") for t in top_terms]
        names[cid] = joiner.join(safe_terms)
    # handle noise cluster -1
    if -1 in names:
        names[-1] = "noise"
    return names


# ============================================================================
# Two-tier categorization: HDBSCAN-driven subcategory discovery with
# deterministic Title Case naming and a Non-Repetitive bucket. Reuses
# vectorize_texts + cluster_texts + get_top_keywords_per_cluster.
# ============================================================================


_TITLE_CASE_STOPWORDS = {
    "a", "an", "and", "or", "of", "the", "to", "for", "in", "on", "at",
    "with", "by", "from", "via", "is", "are", "be", "as",
}
_PUNCT_STRIP = ".,;:!?'\"()[]{}<>"


def _case_form_lookup(texts: List[str]) -> Dict[str, str]:
    """Build a {lowercase_token: most_common_surface_form} map from a corpus.

    Used so generated names preserve original capitalization for technical
    terms — ``"SAP"`` stays ``SAP`` rather than collapsing to ``Sap``, ``RP1``
    stays ``RP1`` not ``Rp1``. Falls back to Title Case for tokens we never
    saw or that appear with inconsistent casing.
    """
    from collections import Counter
    counter: Dict[str, Counter] = {}
    for text in texts:
        if not text:
            continue
        for token in str(text).split():
            stripped = token.strip(_PUNCT_STRIP).strip()
            if not stripped:
                continue
            counter.setdefault(stripped.lower(), Counter())[stripped] += 1
    return {lo: c.most_common(1)[0][0] for lo, c in counter.items()}


def _apply_case_map(token: str, case_map: Dict[str, str]) -> str:
    """Substitute ``token`` with its most-common original surface form.

    Only honors the corpus surface form when it carries non-default casing
    (acronyms like ``SAP``, mixed case like ``PowerBI``). When the corpus uses
    lowercase, fall through to Title Case so ``"failed"`` becomes ``Failed``,
    not ``failed``, in the rendered name. Stop-word connectives stay lower.
    """
    if not token:
        return token
    lo = token.lower()
    if lo in _TITLE_CASE_STOPWORDS:
        return lo
    surface = case_map.get(lo)
    if surface and surface != surface.lower():
        return surface
    # Preserve short all-caps acronyms passed in directly.
    if 2 <= len(token) <= 4 and token.isalpha() and token.isupper():
        return token
    return token[:1].upper() + token[1:].lower()


def phrase_name_from_keywords(
    top_keywords: List[Tuple[str, float]], *, top_n: int = 4, max_chars: int = 60,
    case_map: Optional[Dict[str, str]] = None,
) -> str:
    """Render top TF-IDF (or n-gram TF-IDF) keywords as a human-readable phrase.

    Strategy:
      1. Take the top-1 term. If it already spans 2+ tokens (an n-gram like
         "batch job failure"), Title-Case and return it as the whole name.
      2. Otherwise, concatenate the top-N unigrams in score order, dedupe
         repeats, Title-Case, cap at max_chars.

    When ``case_map`` is provided, each token is replaced with its most-common
    surface form from the source corpus — so ``"SAP"`` stays uppercase rather
    than collapsing to ``"Sap"``. Deterministic given the same keyword list.
    """
    if not top_keywords:
        return ""
    case_map = case_map or {}

    def _title(token: str) -> str:
        return _apply_case_map(token, case_map)

    top_term = str(top_keywords[0][0]).strip()
    if " " in top_term:
        words = [w for w in top_term.split() if w]
        phrase = " ".join(_title(w) for w in words)
        return phrase[:max_chars].rstrip()

    # Otherwise flatten all keywords to individual words and dedupe at the
    # word level — keywords like "rare" and "rare coffee machine" share
    # tokens and the rendered phrase shouldn't repeat them.
    max_words = max(top_n, 5)
    seen: set = set()
    parts: List[str] = []
    for term, _score in top_keywords:
        for word in str(term).strip().split():
            if not word:
                continue
            lo = word.lower()
            if lo in seen:
                continue
            seen.add(lo)
            parts.append(_title(word))
            if len(parts) >= max_words:
                break
        if len(parts) >= max_words:
            break
    phrase = " ".join(parts).rstrip(" -_.,;:/\\")
    return phrase[:max_chars].rstrip()


def _representative_samples(
    X, sub_labels: np.ndarray, centroids: np.ndarray, cluster_ids: List[int],
    texts: List[str], *, k: int = 5,
) -> Dict[int, List[str]]:
    """Return the ``k`` texts closest to each cluster's centroid by cosine.

    Used both for sample-derived name candidates (engine internal) and for the
    Results-tab drill-down (GUI consumes via the TaxonomyResult).
    """
    out: Dict[int, List[str]] = {}
    if centroids.size == 0:
        return {cid: [] for cid in cluster_ids}
    centroid_norms = np.linalg.norm(centroids, axis=1)
    centroid_norms[centroid_norms == 0] = 1.0
    for idx, cid in enumerate(cluster_ids):
        mask = sub_labels == cid
        if not np.any(mask):
            out[cid] = []
            continue
        rows = np.asarray(_to_dense(X[mask]), dtype=np.float64)
        row_norms = np.linalg.norm(rows, axis=1)
        row_norms[row_norms == 0] = 1.0
        sims = (rows @ centroids[idx]) / (row_norms * centroid_norms[idx])
        top_local = np.argsort(-sims)[:k]
        cluster_idx = np.where(mask)[0]
        out[cid] = [str(texts[cluster_idx[i]]) for i in top_local]
    return out


def _sample_derived_name(
    samples: List[str], top_keywords: List[Tuple[str, float]],
    case_map: Dict[str, str], *, min_repetition: int = 2,
) -> Optional[str]:
    """Pick a short phrase from sample tickets that's a better candidate than glued keywords.

    Heuristic: scan 3-5-word n-grams in each sample; require ≥2 token-overlap
    with the cluster's top keywords; require the n-gram to appear in
    ``min_repetition`` distinct samples. Among candidates, prefer most-common
    then shortest. Returns ``None`` if no candidate qualifies — caller falls
    back to ``phrase_name_from_keywords``.
    """
    if len(samples) < min_repetition or not top_keywords:
        return None
    keyword_tokens = {str(kw).lower() for kw, _ in top_keywords}
    if not keyword_tokens:
        return None

    from collections import Counter
    counts: Counter = Counter()
    for sample in samples:
        tokens = [t.strip(_PUNCT_STRIP) for t in str(sample).split()]
        tokens = [t for t in tokens if t]
        seen_in_sample: set = set()
        for n in (3, 4, 5):
            for i in range(len(tokens) - n + 1):
                gram = tokens[i:i + n]
                gram_lo = {t.lower() for t in gram}
                if len(gram_lo & keyword_tokens) < 2:
                    continue
                key = " ".join(gram).lower()
                # Count once per sample even if the n-gram appears twice in one sample.
                if key in seen_in_sample:
                    continue
                seen_in_sample.add(key)
                counts[key] += 1

    if not counts:
        return None
    # Sort: more samples > shorter > earlier (alphabetical fallback for determinism).
    best_key, best_count = max(
        counts.items(), key=lambda x: (x[1], -len(x[0].split()), -ord(x[0][0]))
    )
    if best_count < min_repetition:
        return None
    words = [_apply_case_map(w, case_map) for w in best_key.split()]
    return " ".join(words)[:60]


def _compute_c_tf_idf_keywords(
    texts: List[str], labels: np.ndarray, *,
    ngram_range: Tuple[int, int] = (1, 3),
    max_features: int = 2000,
    top_n: int = 8,
) -> Dict[int, List[Tuple[str, float]]]:
    """Class-based TF-IDF over the cluster mega-documents.

    Standard BERTopic recipe:
      1. Concatenate all texts within a cluster into one big "document".
      2. CountVectorizer over those mega-documents.
      3. Apply TF-IDF transform — IDF here weights terms by how *few*
         clusters they appear in, so the top-scoring terms per cluster are
         the ones that make that cluster distinct from the others.

    Returns a {cluster_id: [(term, score), …]} dict ordered by descending
    score. Cluster id ``-1`` (noise) is excluded.
    """
    cluster_ids = sorted({int(label) for label in np.unique(labels) if label != -1})
    if not cluster_ids:
        return {}

    docs_per_cluster: List[str] = []
    for cid in cluster_ids:
        mask = labels == cid
        joined = " ".join(str(t) for i, t in enumerate(texts) if mask[i])
        docs_per_cluster.append(joined)

    from sklearn.feature_extraction.text import CountVectorizer

    try:
        count_vec = CountVectorizer(
            ngram_range=ngram_range, stop_words="english", max_features=max_features,
        )
        counts = count_vec.fit_transform(docs_per_cluster)
    except ValueError:
        # Vocabulary may be empty if all texts are stop-words / too short.
        return {cid: [] for cid in cluster_ids}

    c_tfidf = TfidfTransformer().fit_transform(counts)
    features = count_vec.get_feature_names_out()

    out: Dict[int, List[Tuple[str, float]]] = {}
    for row_idx, cid in enumerate(cluster_ids):
        scores = np.asarray(c_tfidf[row_idx].todense()).flatten()
        order = scores.argsort()[::-1][:top_n]
        out[cid] = [(str(features[i]), float(scores[i])) for i in order if scores[i] > 0]
    return out


def _subclusters_at_scale(
    X, *, min_cluster_size: int, random_state: int = 42,
) -> np.ndarray:
    """MiniBatchKMeans + percentile-based noise tagging for large corpora.

    HDBSCAN's worst-case runtime gets impractical above ~25k rows. This is the
    fallback path: pick ``k`` heuristically (~one cluster per 200 rows, capped
    8-60), run MiniBatchKMeans, then tag rows in the top 5% of centroid-
    distance as noise (`-1`) so the downstream Non-Repetitive bucket still
    captures the long tail.
    """
    X_dense = np.asarray(_to_dense(X), dtype=np.float64)
    n_rows = X_dense.shape[0]
    k = max(8, min(60, n_rows // 200))
    model = MiniBatchKMeans(
        n_clusters=k, init="k-means++", n_init=10, random_state=random_state,
    )
    labels = model.fit_predict(X_dense).astype(np.int64)

    centroids = model.cluster_centers_
    per_row_dist = np.zeros(n_rows, dtype=np.float64)
    for i in range(k):
        mask = labels == i
        if not np.any(mask):
            continue
        per_row_dist[mask] = np.linalg.norm(X_dense[mask] - centroids[i], axis=1)
    p95 = float(np.percentile(per_row_dist, 95))
    labels[per_row_dist > p95] = -1
    # Apply the same min_cluster_size cutoff to keep semantics consistent.
    unique, counts = np.unique(labels, return_counts=True)
    for cid, count in zip(unique, counts):
        if cid == -1:
            continue
        if count < min_cluster_size:
            labels[labels == cid] = -1
    return labels


def _apply_umap_reduction(X, *, n_components: int = 15, random_state: int = 42):
    """Optional UMAP dim reduction — only used when umap-learn is installed.

    The BERTopic recipe for embedding-mode clustering. Drops 384-dim sentence
    embeddings to ~15 dims, which kills high-D distance concentration and
    produces tighter, finer-grained HDBSCAN clusters. Silently passes X
    through unchanged when umap-learn isn't available so the caller doesn't
    need to special-case the missing dep.
    """
    if not _UMAP_AVAILABLE:
        return X, False
    try:
        import umap as _umap  # type: ignore[import-untyped]
    except Exception:
        return X, False
    X_dense = np.asarray(_to_dense(X), dtype=np.float32)
    n_components = min(n_components, max(2, X_dense.shape[0] - 2))
    # n_neighbors is bounded by sample count - 1.
    n_neighbors = min(15, max(2, X_dense.shape[0] - 1))
    try:
        reducer = _umap.UMAP(
            n_components=n_components,
            n_neighbors=n_neighbors,
            metric="cosine",
            random_state=random_state,
        )
        return reducer.fit_transform(X_dense), True
    except Exception:
        return X, False


def _subcluster_fingerprint(top_keywords: List[Tuple[str, float]], *, n: int = 8) -> str:
    """Stable short hash of a subcluster's top keywords.

    Sorted so the fingerprint survives reordering. Truncated to 16 hex chars.
    Used to preserve user-edited names across re-runs even when cleaning
    settings shift the cluster id.
    """
    tokens = sorted(str(t).strip().lower() for t, _ in top_keywords[:n] if t)
    payload = "|".join(tokens).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()[:16]


def _mark_non_repetitive(labels: np.ndarray, min_size: int) -> np.ndarray:
    """Relabel HDBSCAN noise (-1) plus any cluster with `< min_size` rows as -1.

    HDBSCAN already emits -1 for noise. This helper additionally pushes tiny
    surviving clusters into the same -1 bucket so the downstream `Non-Repetitive`
    label captures both signals.
    """
    out = np.asarray(labels, dtype=np.int64).copy()
    unique, counts = np.unique(out, return_counts=True)
    for label, count in zip(unique, counts):
        if label == -1:
            continue
        if count < min_size:
            out[out == label] = -1
    return out


def _compute_centroids(X, sub_labels: np.ndarray) -> Tuple[np.ndarray, List[int]]:
    """Compute one mean vector per non-noise subcluster.

    Returns (centroids, cluster_ids) where ``centroids[i]`` corresponds to
    ``cluster_ids[i]``. For sparse X (TF-IDF), centroids come out dense (the
    matrices are small — one row per subcluster).
    """
    X_dense = _to_dense(X)
    ids = sorted(int(i) for i in np.unique(sub_labels) if i != -1)
    if not ids:
        # Match X's column count even if no clusters survived, so caller doesn't
        # need to special-case the shape.
        return np.zeros((0, X_dense.shape[1] if X_dense.ndim == 2 else 0)), []
    centroids = np.vstack([
        np.asarray(X_dense[sub_labels == cid]).mean(axis=0) for cid in ids
    ])
    return centroids, ids


def _confidence_scores(
    X, sub_labels: np.ndarray, centroids: np.ndarray, cluster_ids: List[int],
    *,
    hdbscan_probabilities: Optional[np.ndarray] = None,
    prob_weight: float = 0.6, margin_weight: float = 0.4,
) -> np.ndarray:
    """Per-row confidence score combining multiple calibrated signals.

    Components (all in [0, 1]):
      • ``cosine_best``     — cosine to assigned centroid (the original signal)
      • ``cosine_margin``   — best cosine minus second-best across all centroids;
                              captures *ambiguity* (ticket sitting between two
                              clusters has small margin even with high cosine)
      • ``hdbscan_prob``    — when ``hdbscan_probabilities`` is supplied (only
                              from ``categorize_taxonomy``, not ``apply_taxonomy``),
                              HDBSCAN's per-point membership probability blended in.

    Combination:
      • With HDBSCAN probs:   prob_weight * prob + margin_weight * margin
      • Fallback (no probs):  0.7 * cosine_best + 0.3 * margin

    Non-Repetitive rows (`sub_labels == -1`) are forced to 0.0.
    """
    n_rows = int(np.asarray(sub_labels).shape[0])
    out = np.zeros(n_rows, dtype=np.float64)
    if centroids.size == 0:
        return out

    id_to_idx = {cid: i for i, cid in enumerate(cluster_ids)}
    centroid_norms = np.linalg.norm(centroids, axis=1)
    centroid_norms[centroid_norms == 0] = 1.0
    centroids_unit = centroids / centroid_norms[:, None]

    X_dense = np.asarray(_to_dense(X), dtype=np.float64)
    row_norms = np.linalg.norm(X_dense, axis=1)
    row_norms[row_norms == 0] = 1.0
    # Full similarity matrix: rows × centroids. For 384-dim embeddings + a few
    # hundred centroids this is small enough to materialize.
    sims_all = (X_dense @ centroids_unit.T) / row_norms[:, None]
    sims_all = np.clip(sims_all, 0.0, 1.0)

    for cid, idx in id_to_idx.items():
        mask = sub_labels == cid
        if not np.any(mask):
            continue
        cosine_best = sims_all[mask, idx]
        if sims_all.shape[1] > 1:
            # Margin = best − second-best. Mask the assigned column to -inf
            # then take row-max.
            others = sims_all[mask].copy()
            others[:, idx] = -np.inf
            cosine_second = others.max(axis=1)
            margin = np.clip(cosine_best - cosine_second, 0.0, 1.0)
        else:
            margin = np.ones_like(cosine_best)  # only one cluster → no ambiguity

        if hdbscan_probabilities is not None:
            probs = np.asarray(hdbscan_probabilities)[mask].astype(np.float64)
            score = prob_weight * probs + margin_weight * margin
        else:
            score = 0.7 * cosine_best + 0.3 * margin
        out[mask] = np.clip(score, 0.0, 1.0)
    return out


@dataclass
class TaxonomyResult:
    """Outcome of categorize_taxonomy() / apply_taxonomy().

    Three per-row arrays (`repetitive`, `subcategory`, `confidence`) are aligned
    with the input texts. The remaining fields support Save Taxonomy + re-runs,
    plus Results-tab drill-down (`samples_by_cluster`) and audit metadata
    (`manifest`).
    """

    repetitive: List[str]
    subcategory: List[str]
    confidence: List[float]
    subcluster_labels: np.ndarray
    subcategory_names: Dict[int, str]
    sub_centroids: np.ndarray
    sub_fingerprints: Dict[int, str]
    vectorizer: Any
    stats: Dict[str, Any]
    # Optional enrichments (populated by categorize_taxonomy; may be empty for
    # legacy callers or fast-path apply_taxonomy outputs).
    samples_by_cluster: Dict[int, List[str]] = field(default_factory=dict)
    avg_confidence_by_cluster: Dict[int, float] = field(default_factory=dict)
    manifest: Dict[str, Any] = field(default_factory=dict)


_REPETITIVE = "Repetitive"
_NON_REPETITIVE = "Non-Repetitive"


def _disambiguate_duplicate_names(names: Dict[int, str]) -> Dict[int, str]:
    """Deterministic `name (1)`, `name (2)` suffixes for colliding names.

    Iteration order is by ascending cluster id so the suffix assignment is
    reproducible across runs.
    """
    seen: Dict[str, int] = {}
    out: Dict[int, str] = {}
    for cid in sorted(names.keys()):
        base = names[cid] or f"Subcluster {cid}"
        count = seen.get(base, 0)
        out[cid] = base if count == 0 else f"{base} ({count})"
        seen[base] = count + 1
    return out


def categorize_taxonomy(
    texts: List[str],
    *,
    precomputed: Optional[Tuple[Any, Any]] = None,
    vectorizer_kind: str = "embedding",
    min_cluster_size: int = 5,
    min_samples: int = 3,
    non_repetitive_min_size: int = 5,
    name_ngram_range: Tuple[int, int] = (1, 3),
    user_renames: Optional[Dict[str, str]] = None,
    use_umap: Optional[bool] = None,
    n_samples_per_cluster: int = 5,
    progress_cb: Optional[Callable[[int, str], None]] = None,
) -> TaxonomyResult:
    """Discover a single-level subcategory taxonomy from cleaned texts.

    Pipeline:
      1. Vectorize (embedding or TF-IDF).
      2. (optional) UMAP dim-reduction when embedding-mode + umap-learn installed.
      3. Sub-cluster: HDBSCAN normally; MiniBatchKMeans fallback above
         ``_CATEGORIZATION_SCALE_GUARD`` rows.
      4. Mark noise + tiny clusters as Non-Repetitive.
      5. Name each cluster via c-TF-IDF + sample-derived candidates +
         original-case preservation. Deterministic.
      6. Score confidence by combining HDBSCAN probabilities (when available)
         with cosine + margin-to-next-best-cluster.

    Re-runs with the same inputs + ``user_renames`` produce identical output.

    Set ``use_umap=False`` to disable UMAP even when installed; the default
    ``None`` means "use UMAP if embedding-mode + umap-learn installed".
    """
    user_renames = dict(user_renames or {})
    start_ts = datetime.now(timezone.utc)

    def _emit(pct: int, msg: str) -> None:
        if progress_cb is not None:
            try:
                progress_cb(pct, msg)
            except Exception:
                pass

    _emit(5, "Vectorizing…")
    if precomputed is not None:
        vec, X = precomputed
    else:
        kind = vectorizer_kind
        if kind == "embedding" and not _ST_AVAILABLE:
            warnings.warn(
                "sentence-transformers not installed; falling back to TF-IDF "
                "for categorization. Install with: pip install sentence-transformers"
            )
            kind = "tfidf"
        vec, X = vectorize_texts(texts, vectorizer_kind=kind)

    is_embedding = isinstance(vec, EmbeddingVectorizer)

    # Optional UMAP reduction — only meaningful on dense embedding vectors.
    # The reduced matrix is used purely as input to HDBSCAN; the unreduced X
    # is kept around for centroid + confidence calculations so re-applications
    # of the taxonomy don't need UMAP at apply time.
    umap_applied = False
    X_for_clustering = X
    if is_embedding and (use_umap is True or (use_umap is None and _UMAP_AVAILABLE)):
        _emit(25, "Reducing dimensionality (UMAP)…")
        X_for_clustering, umap_applied = _apply_umap_reduction(X)

    _emit(40, "Sub-clustering…")
    n_rows = len(texts)
    hdbscan_probs: Optional[np.ndarray] = None
    scale_path: str
    if n_rows >= _CATEGORIZATION_SCALE_GUARD:
        # Big-data fallback — MiniBatchKMeans + percentile noise tagging.
        warnings.warn(
            f"Corpus size {n_rows} ≥ {_CATEGORIZATION_SCALE_GUARD}: using "
            "MiniBatchKMeans + percentile noise tagging in place of HDBSCAN."
        )
        sub_raw = _subclusters_at_scale(X_for_clustering, min_cluster_size=min_cluster_size)
        scale_path = "minibatch_kmeans_scale_guard"
    else:
        if not _HDBSCAN_AVAILABLE:
            raise ImportError(
                "HDBSCAN is required for categorization. Install with: pip install hdbscan"
            )
        X_dense_for_cluster = _to_dense(X_for_clustering)
        model = hdbscan.HDBSCAN(
            min_cluster_size=min_cluster_size,
            min_samples=min_samples,
            metric="euclidean",
            cluster_selection_method="eom",
            prediction_data=True,
        )
        sub_raw = model.fit_predict(X_dense_for_cluster)
        hdbscan_probs = np.asarray(getattr(model, "probabilities_", []), dtype=np.float64)
        if hdbscan_probs.shape[0] != n_rows:
            hdbscan_probs = None  # safety: stay silent if shapes don't line up
        scale_path = "hdbscan"

    sub_labels = _mark_non_repetitive(sub_raw, non_repetitive_min_size)
    # Centroids are computed in the original (unreduced) X space so they are
    # directly comparable to encoded fresh data in apply_taxonomy.
    centroids, cluster_ids = _compute_centroids(X, sub_labels)

    _emit(70, "Naming subcategories…")
    # Case-preservation map built from the source corpus so generated names
    # keep "SAP", "RP1", "ECC" rather than collapsing them to Title Case.
    case_map = _case_form_lookup(texts)
    # c-TF-IDF over cluster mega-documents — names reflect what's *distinctive*
    # about each cluster, not what's merely frequent inside it.
    top_kw_for_naming = _compute_c_tf_idf_keywords(
        texts, sub_labels, ngram_range=name_ngram_range,
    )
    samples_by_cluster = _representative_samples(
        X, sub_labels, centroids, cluster_ids, texts, k=max(n_samples_per_cluster, 5),
    )

    raw_names: Dict[int, str] = {}
    fingerprints: Dict[int, str] = {}
    for cid in cluster_ids:
        kw_list = top_kw_for_naming.get(cid, [])
        fingerprints[cid] = _subcluster_fingerprint(kw_list)
        # User-edited names win, keyed on fingerprint so cleaning tweaks
        # that re-id the cluster still preserve the rename.
        if fingerprints[cid] in user_renames:
            raw_names[cid] = user_renames[fingerprints[cid]]
            continue
        # Prefer a sample-derived candidate when one qualifies — it reads
        # like a real ticket subject rather than glued keywords.
        sample_candidate = _sample_derived_name(
            samples_by_cluster.get(cid, []), kw_list, case_map,
        )
        if sample_candidate:
            raw_names[cid] = sample_candidate
            continue
        phrase = phrase_name_from_keywords(kw_list, case_map=case_map)
        raw_names[cid] = phrase or f"Subcluster {cid}"
    subcategory_names = _disambiguate_duplicate_names(raw_names)

    _emit(90, "Confidence scoring…")
    confidence = _confidence_scores(
        X, sub_labels, centroids, cluster_ids,
        hdbscan_probabilities=hdbscan_probs,
    )

    avg_conf_by_cluster: Dict[int, float] = {}
    for cid in cluster_ids:
        mask = sub_labels == cid
        if np.any(mask):
            avg_conf_by_cluster[cid] = float(np.mean(confidence[mask]))

    repetitive: List[str] = []
    subcategory: List[str] = []
    for label in sub_labels.tolist():
        if label == -1:
            repetitive.append(_NON_REPETITIVE)
            subcategory.append(_NON_REPETITIVE)
        else:
            repetitive.append(_REPETITIVE)
            subcategory.append(subcategory_names.get(int(label), f"Subcluster {label}"))

    n_non_rep = int(np.sum(sub_labels == -1))
    stats = {
        "n_subclusters": len(cluster_ids),
        "n_non_repetitive": n_non_rep,
        "pct_non_repetitive": (n_non_rep / n_rows) if n_rows else 0.0,
        "vectorizer_kind": "embedding" if is_embedding else "tfidf",
        "umap_applied": bool(umap_applied),
        "scale_path": scale_path,
    }
    manifest = {
        "created_at": start_ts.isoformat(),
        "vectorizer_kind": stats["vectorizer_kind"],
        "embedding_model": getattr(vec, "model_name", None) if is_embedding else None,
        "min_cluster_size": int(min_cluster_size),
        "min_samples": int(min_samples),
        "non_repetitive_min_size": int(non_repetitive_min_size),
        "name_ngram_range": list(name_ngram_range),
        "umap_applied": bool(umap_applied),
        "scale_path": scale_path,
        "n_rows": int(n_rows),
        "n_subclusters": stats["n_subclusters"],
    }
    _emit(100, "Done")
    return TaxonomyResult(
        repetitive=repetitive,
        subcategory=subcategory,
        confidence=[float(c) for c in confidence.tolist()],
        subcluster_labels=sub_labels,
        subcategory_names=subcategory_names,
        sub_centroids=centroids,
        sub_fingerprints=fingerprints,
        samples_by_cluster=samples_by_cluster,
        avg_confidence_by_cluster=avg_conf_by_cluster,
        manifest=manifest,
        vectorizer=vec,
        stats=stats,
    )


def apply_taxonomy(
    texts: List[str], taxonomy: Dict[str, Any], *, confidence_threshold: float = 0.45,
) -> TaxonomyResult:
    """Apply a saved taxonomy to fresh texts.

    Uses the saved vectorizer to encode the new texts, then assigns each row to
    the nearest subcluster centroid (cosine). Rows below ``confidence_threshold``
    fall into Non-Repetitive. Bypasses HDBSCAN entirely — fast path for batch
    re-runs against a previously-trained taxonomy.

    Expects a payload produced by IOService.save_taxonomy.
    """
    required = ("vectorizer", "sub_centroids", "subcategory_names")
    for key in required:
        if key not in taxonomy:
            raise RuntimeError(f"Taxonomy payload missing required key: {key!r}")

    vec = taxonomy["vectorizer"]
    centroids = np.asarray(taxonomy["sub_centroids"])
    subcategory_names: Dict[int, str] = {
        int(k): str(v) for k, v in taxonomy["subcategory_names"].items()
    }
    fingerprints: Dict[int, str] = {
        int(k): str(v) for k, v in (taxonomy.get("sub_fingerprints") or {}).items()
    }
    # Apply user_renames on top of the stored names so renames survive
    # save → load → re-apply.
    for cid, fp in fingerprints.items():
        if fp in (taxonomy.get("user_renames") or {}):
            subcategory_names[cid] = taxonomy["user_renames"][fp]

    if not hasattr(vec, "transform"):
        raise RuntimeError("Saved vectorizer has no transform() — cannot apply taxonomy.")
    X_new = vec.transform(list(texts))
    X_dense = _to_dense(X_new)

    if centroids.size == 0:
        labels = np.full(len(texts), -1, dtype=np.int64)
        confidence = np.zeros(len(texts), dtype=np.float64)
        cluster_ids = sorted(subcategory_names.keys())
    else:
        centroid_norms = np.linalg.norm(centroids, axis=1)
        centroid_norms[centroid_norms == 0] = 1.0
        centroids_unit = centroids / centroid_norms[:, None]
        rows = np.asarray(X_dense, dtype=np.float64)
        row_norms = np.linalg.norm(rows, axis=1)
        row_norms[row_norms == 0] = 1.0
        sims = (rows @ centroids_unit.T) / row_norms[:, None]
        sims = np.clip(sims, 0.0, 1.0)
        best_idx = sims.argmax(axis=1)
        best_sim = sims[np.arange(sims.shape[0]), best_idx]
        cluster_ids = sorted(subcategory_names.keys())
        labels = np.array([cluster_ids[i] for i in best_idx], dtype=np.int64)
        labels[best_sim < confidence_threshold] = -1
        # Blend best cosine with margin-to-next-best for a more calibrated
        # confidence — same recipe as the from-scratch path, minus HDBSCAN probs.
        confidence = _confidence_scores(X_new, labels, centroids, cluster_ids)
        confidence[labels == -1] = 0.0

    repetitive: List[str] = []
    subcategory: List[str] = []
    for label in labels.tolist():
        if label == -1:
            repetitive.append(_NON_REPETITIVE)
            subcategory.append(_NON_REPETITIVE)
        else:
            repetitive.append(_REPETITIVE)
            subcategory.append(subcategory_names.get(int(label), f"Subcluster {label}"))

    n_rows = len(texts)
    n_non_rep = int(np.sum(labels == -1))
    stats = {
        "n_subclusters": len(subcategory_names),
        "n_non_repetitive": n_non_rep,
        "pct_non_repetitive": (n_non_rep / n_rows) if n_rows else 0.0,
        "applied_from_saved": True,
        "confidence_threshold": confidence_threshold,
    }
    # Build samples + avg confidence per cluster from the loaded centroids
    # so the Results tab drill-down works on the fast-apply path too.
    samples_by_cluster = _representative_samples(
        X_new, labels, centroids, cluster_ids, texts, k=5,
    ) if centroids.size else {}
    avg_conf_by_cluster: Dict[int, float] = {}
    for cid in cluster_ids:
        mask = labels == cid
        if np.any(mask):
            avg_conf_by_cluster[int(cid)] = float(np.mean(confidence[mask]))
    # Preserve the source manifest if present; mark this run as an application.
    manifest = dict(taxonomy.get("manifest") or {})
    manifest["applied_at"] = datetime.now(timezone.utc).isoformat()
    manifest["applied_to_n_rows"] = int(n_rows)
    manifest["confidence_threshold"] = float(confidence_threshold)
    return TaxonomyResult(
        repetitive=repetitive,
        subcategory=subcategory,
        confidence=[float(c) for c in confidence.tolist()],
        subcluster_labels=labels,
        subcategory_names=subcategory_names,
        sub_centroids=centroids,
        sub_fingerprints=fingerprints,
        samples_by_cluster=samples_by_cluster,
        avg_confidence_by_cluster=avg_conf_by_cluster,
        manifest=manifest,
        vectorizer=vec,
        stats=stats,
    )


def _rebuild_taxonomy_after_labels_change(
    result: TaxonomyResult,
    X,
    texts: List[str],
    new_sub_labels: np.ndarray,
    *,
    name_ngram_range: Tuple[int, int] = (1, 3),
    user_renames: Optional[Dict[str, str]] = None,
) -> TaxonomyResult:
    """Recompute names + centroids + confidence after a manual label edit.

    Shared internals for ``merge_clusters`` and ``split_cluster``: both produce
    a new label array and need the rest of the TaxonomyResult re-derived from
    it. Preserves the caller's vectorizer, manifest origin, and respects any
    user_renames keyed on fingerprint.
    """
    user_renames = dict(user_renames or {})
    centroids, cluster_ids = _compute_centroids(X, new_sub_labels)
    case_map = _case_form_lookup(texts)
    top_kw = _compute_c_tf_idf_keywords(texts, new_sub_labels, ngram_range=name_ngram_range)
    samples_by_cluster = _representative_samples(
        X, new_sub_labels, centroids, cluster_ids, texts, k=5,
    )

    raw_names: Dict[int, str] = {}
    fingerprints: Dict[int, str] = {}
    for cid in cluster_ids:
        kw_list = top_kw.get(cid, [])
        fingerprints[cid] = _subcluster_fingerprint(kw_list)
        if fingerprints[cid] in user_renames:
            raw_names[cid] = user_renames[fingerprints[cid]]
            continue
        sample_candidate = _sample_derived_name(
            samples_by_cluster.get(cid, []), kw_list, case_map,
        )
        if sample_candidate:
            raw_names[cid] = sample_candidate
            continue
        raw_names[cid] = phrase_name_from_keywords(kw_list, case_map=case_map) or f"Subcluster {cid}"
    subcategory_names = _disambiguate_duplicate_names(raw_names)

    confidence = _confidence_scores(X, new_sub_labels, centroids, cluster_ids)
    avg_conf_by_cluster: Dict[int, float] = {}
    for cid in cluster_ids:
        mask = new_sub_labels == cid
        if np.any(mask):
            avg_conf_by_cluster[cid] = float(np.mean(confidence[mask]))

    repetitive: List[str] = []
    subcategory: List[str] = []
    for label in new_sub_labels.tolist():
        if label == -1:
            repetitive.append(_NON_REPETITIVE)
            subcategory.append(_NON_REPETITIVE)
        else:
            repetitive.append(_REPETITIVE)
            subcategory.append(subcategory_names.get(int(label), f"Subcluster {label}"))

    n_rows = len(texts)
    n_non_rep = int(np.sum(new_sub_labels == -1))
    stats = dict(result.stats)
    stats.update({
        "n_subclusters": len(cluster_ids),
        "n_non_repetitive": n_non_rep,
        "pct_non_repetitive": (n_non_rep / n_rows) if n_rows else 0.0,
        "post_edit": True,
    })
    manifest = dict(result.manifest)
    manifest["last_edit_at"] = datetime.now(timezone.utc).isoformat()
    manifest["n_subclusters"] = len(cluster_ids)

    return TaxonomyResult(
        repetitive=repetitive,
        subcategory=subcategory,
        confidence=[float(c) for c in confidence.tolist()],
        subcluster_labels=new_sub_labels,
        subcategory_names=subcategory_names,
        sub_centroids=centroids,
        sub_fingerprints=fingerprints,
        samples_by_cluster=samples_by_cluster,
        avg_confidence_by_cluster=avg_conf_by_cluster,
        manifest=manifest,
        vectorizer=result.vectorizer,
        stats=stats,
    )


def merge_clusters(
    result: TaxonomyResult,
    X,
    texts: List[str],
    cluster_ids: List[int],
    *,
    user_renames: Optional[Dict[str, str]] = None,
) -> TaxonomyResult:
    """Merge two or more clusters into the lowest-id of the set.

    Recomputes names + centroids + confidence for the merged cluster (and
    leaves the others alone). Used by the Results-tab "Merge into…" action
    when the user spots two clusters that should be one.
    """
    if len(cluster_ids) < 2:
        raise ValueError("merge_clusters requires at least 2 cluster ids.")
    target = min(cluster_ids)
    new_labels = np.asarray(result.subcluster_labels, dtype=np.int64).copy()
    for cid in cluster_ids:
        if cid == target:
            continue
        new_labels[new_labels == cid] = target
    return _rebuild_taxonomy_after_labels_change(
        result, X, texts, new_labels, user_renames=user_renames,
    )


def split_cluster(
    result: TaxonomyResult,
    X,
    texts: List[str],
    cluster_id: int,
    *,
    k: int = 2,
    random_state: int = 42,
    user_renames: Optional[Dict[str, str]] = None,
) -> TaxonomyResult:
    """Split a single cluster into ``k`` sub-clusters via k-means.

    Runs k-means on the cluster's vectors, then assigns the new groups fresh
    cluster ids (`max(existing) + 1`, `+ 2`, …). The original cluster id is
    consumed. Other clusters and the Non-Repetitive bucket are untouched.
    """
    if k < 2:
        raise ValueError("split_cluster requires k >= 2.")
    new_labels = np.asarray(result.subcluster_labels, dtype=np.int64).copy()
    mask = new_labels == cluster_id
    if not np.any(mask):
        raise ValueError(f"Cluster id {cluster_id} not present in subcluster_labels.")
    member_count = int(np.sum(mask))
    if member_count < k:
        raise ValueError(
            f"Cluster {cluster_id} has only {member_count} rows — cannot split into {k}."
        )
    X_member = np.asarray(_to_dense(X[mask]), dtype=np.float64)
    model = KMeans(n_clusters=k, init="k-means++", n_init=10, random_state=random_state)
    sub = model.fit_predict(X_member)
    next_id = int(new_labels.max()) + 1
    # First k-means group reuses the original cluster id; remaining groups get fresh ids.
    idx_in_labels = np.where(mask)[0]
    for j in range(k):
        sub_mask_in_member = sub == j
        new_id = cluster_id if j == 0 else next_id + (j - 1)
        new_labels[idx_in_labels[sub_mask_in_member]] = new_id
    return _rebuild_taxonomy_after_labels_change(
        result, X, texts, new_labels, user_renames=user_renames,
    )


def compute_embedding_2d(X, method: str = "pca", perplexity: int = 30, random_state: int = 42) -> np.ndarray:
    """Return a 2D embedding of X using PCA or t-SNE. Reusable by GUI for inline rendering."""
    dense = _to_dense(X)
    if method == "pca":
        reducer = PCA(n_components=2, random_state=random_state)
    elif method == "tsne":
        effective_perplexity = max(2, min(perplexity, max(2, dense.shape[0] - 1)))
        reducer = TSNE(n_components=2, perplexity=effective_perplexity, random_state=random_state)
    else:
        raise ValueError("Unknown visualization method: choose 'pca' or 'tsne'")
    return reducer.fit_transform(dense)


def visualize_embeddings(X, labels: np.ndarray, method: str = "pca", perplexity: int = 30, random_state: int = 42, out_path: Optional[str] = None):
    # matplotlib + seaborn imported lazily here so the engine module's cold
    # start doesn't pay for them when the user never opens a CLI viz / GUI plot.
    import matplotlib.pyplot as plt
    import seaborn as sns

    emb = compute_embedding_2d(X, method=method, perplexity=perplexity, random_state=random_state)

    df_vis = pd.DataFrame({"x": emb[:, 0], "y": emb[:, 1], "label": labels})
    plt.figure(figsize=(8, 6))
    palette = sns.color_palette("hsv", len(np.unique(labels)))
    sns.scatterplot(data=df_vis, x="x", y="y", hue="label", palette=palette, legend="full", s=40)
    plt.title(f"Cluster visualization ({method})")
    plt.tight_layout()
    if out_path:
        plt.savefig(out_path)
        print(f"Saved visualization to: {out_path}")
    else:
        plt.show()
    plt.close()


def save_results(df: pd.DataFrame, out_path: str) -> str:
    ext = get_file_extension(out_path)

    def _write(target_path: str):
        if ext == ".xlsx":
            df.to_excel(target_path, index=False, engine="openpyxl")
            return
        if ext == ".xls":
            raise ValueError("Writing .xls output is not supported. Please save as .xlsx, .csv, or .json.")
        if ext == ".csv":
            df.to_csv(target_path, index=False)
            return
        if ext == ".json":
            df.to_json(target_path, orient="records", indent=2, force_ascii=False)
            return
        raise ValueError(
            f"Unsupported output type '{ext or '(none)'}'. Supported types: Excel (.xlsx), CSV (.csv), JSON (.json)."
        )

    try:
        _write(out_path)
        print(f"Saved results to {out_path}")
        return out_path
    except PermissionError:
        base, original_ext = os.path.splitext(out_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt_path = f"{base}_writable_{timestamp}{original_ext}"
        _write(alt_path)
        print(f"Could not overwrite {out_path} (permission denied). Saved to {alt_path} instead.")
        return alt_path


def save_results_excel(df: pd.DataFrame, out_path: str):
    return save_results(df, out_path)


def main():
    parser = argparse.ArgumentParser(description="Cluster text data from Excel, CSV, or JSON files and write cluster labels back to a table file.")
    parser.add_argument("--input", "-i", required=True, help="Input file path (.xlsx, .xls, .csv, .json)")
    parser.add_argument("--sheet", "-s", default=None, help="Sheet name or index for Excel files (optional)")
    parser.add_argument("--column", "-c", required=True, help="Text column name to cluster")
    parser.add_argument("--algorithm", "-a", choices=["kmeans", "dbscan", "agglomerative", "hdbscan"], default="kmeans")
    parser.add_argument("--n_clusters", "-k", type=int, default=5, help="Number of clusters (for kmeans/agglomerative)")
    parser.add_argument("--eps", type=float, default=0.5, help="DBSCAN eps parameter")
    parser.add_argument("--min_samples", type=int, default=5, help="DBSCAN/HDBSCAN min_samples")
    parser.add_argument("--min_cluster_size", type=int, default=5, help="HDBSCAN min_cluster_size")
    parser.add_argument("--max_features", type=int, default=2000, help="Max features for TF-IDF")
    parser.add_argument("--use_hashing", action="store_true", help="Use HashingVectorizer + chunked TF-IDF (memory-efficient; no feature names)")
    parser.add_argument("--chunk_size", type=int, default=10000, help="Chunk size when --use_hashing is on")
    parser.add_argument(
        "--vectorizer",
        choices=["tfidf", "embedding"],
        default="tfidf",
        help="Vectorization mode: 'tfidf' (default) or 'embedding' (requires sentence-transformers)",
    )
    parser.add_argument(
        "--embedding_model",
        default=DEFAULT_EMBEDDING_MODEL,
        help="HuggingFace model name for --vectorizer embedding (default: %(default)s)",
    )
    parser.add_argument(
        "--embedding_device",
        default="cpu",
        help="Device for embedding inference: 'cpu' (default) or 'cuda'",
    )
    parser.add_argument(
        "--embedding_batch_size",
        type=int,
        default=32,
        help="Batch size for embedding inference (default: 32)",
    )
    parser.add_argument("--use_gpu", action="store_true", help="Use GPU-accelerated KMeans via cuML if installed")
    parser.add_argument("--minibatch_threshold", type=int, default=10000, help="Switch to MiniBatchKMeans when n_samples >= this")
    parser.add_argument("--output", "-o", default=None, help="Output path (optional) — defaults to input + _clustered with a matching supported extension")
    parser.add_argument("--visualize", "-v", action="store_true", help="Visualize clusters (PCA)")
    parser.add_argument("--vis_method", choices=["pca", "tsne"], default="pca", help="Visualization method")
    parser.add_argument("--top_n", type=int, default=10, help="Top keywords per cluster")
    parser.add_argument("--name_top_n", type=int, default=3, help="Number of top keywords to form cluster name")
    parser.add_argument("--name_joiner", type=str, default=", ", help="String to join keywords when forming cluster name")
    parser.add_argument("--save_model", action="store_true", help="Save trained clustering model (joblib)")
    parser.add_argument("--load_model", action="store_true", help="Load a saved model (joblib) and apply it to the input instead of training")
    parser.add_argument("--model_path", default=None, help="Path to save (or load with --load_model) the model")
    # Categorization workflow (single-level subcategory discovery).
    parser.add_argument(
        "--categorize",
        action="store_true",
        help="Run single-level subcategory discovery via HDBSCAN with deterministic keyword phrasing. "
             "Output gains 3 columns (Repetitive/Non-Repetitive, Subcategory, Confidence) and, for .xlsx, a pivot sheet.",
    )
    parser.add_argument(
        "--non_repetitive_min_size",
        type=int, default=5,
        help="Categorization: clusters with fewer rows than this are merged into Non-Repetitive (default 5).",
    )
    parser.add_argument(
        "--save_taxonomy", default=None,
        help="Path to save the discovered taxonomy as a .joblib (categorization mode only).",
    )
    parser.add_argument(
        "--load_taxonomy", default=None,
        help="Path to a saved taxonomy .joblib. Applies it to --input via apply_taxonomy "
             "(skips HDBSCAN). Mutually exclusive with --categorize.",
    )
    parser.add_argument(
        "--confidence_threshold", type=float, default=0.45,
        help="Cosine cutoff for --load_taxonomy: rows below this become Non-Repetitive (default 0.45).",
    )

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        raise FileNotFoundError(f"Input file not found: {args.input}")

    df = load_table(args.input, sheet_name=args.sheet)
    if args.column not in df.columns:
        raise ValueError(f"Column '{args.column}' not found in input data. Available columns: {list(df.columns)}")

    text_series = coerce_text_column(df[args.column])
    processed = preprocess_texts(text_series.tolist())

    if args.categorize and args.load_taxonomy:
        raise ValueError("--categorize and --load_taxonomy are mutually exclusive.")

    # Apply a saved taxonomy: encode → nearest centroid → 3-column output + pivot.
    if args.load_taxonomy:
        if not os.path.isfile(args.load_taxonomy):
            raise FileNotFoundError(f"Taxonomy file not found: {args.load_taxonomy}")
        payload = joblib.load(args.load_taxonomy)
        result = apply_taxonomy(
            processed, payload, confidence_threshold=args.confidence_threshold,
        )
        df["Repetitive/Non-Repetitive"] = result.repetitive
        df["Subcategory"] = result.subcategory
        df["Confidence"] = result.confidence
        input_ext = get_file_extension(args.input)
        default_ext = input_ext if input_ext in {".xlsx", ".csv", ".json"} else ".xlsx"
        out_path = args.output or os.path.splitext(args.input)[0] + "_categorized_from_taxonomy" + default_ext
        if out_path.lower().endswith(".xlsx"):
            from textanalyzer.services.io import IOService  # lazy import; avoids cycle
            IOService.save_results_with_pivot(
                df, out_path, sheet_name=args.sheet or "Inc", taxonomy_result=result,
            )
        else:
            save_results(df, out_path)
        stats = result.stats
        print(
            f"Applied taxonomy: {stats['n_subclusters']} subcategories, "
            f"{stats['pct_non_repetitive']:.0%} Non-Repetitive. Saved to: {out_path}"
        )
        return

    # Run categorization from scratch on the loaded data.
    if args.categorize:
        result = categorize_taxonomy(
            processed,
            vectorizer_kind=args.vectorizer,
            min_cluster_size=args.min_cluster_size,
            min_samples=args.min_samples,
            non_repetitive_min_size=args.non_repetitive_min_size,
        )
        df["Repetitive/Non-Repetitive"] = result.repetitive
        df["Subcategory"] = result.subcategory
        df["Confidence"] = result.confidence
        input_ext = get_file_extension(args.input)
        default_ext = input_ext if input_ext in {".xlsx", ".csv", ".json"} else ".xlsx"
        out_path = args.output or os.path.splitext(args.input)[0] + "_categorized" + default_ext
        if out_path.lower().endswith(".xlsx"):
            from textanalyzer.services.io import IOService
            IOService.save_results_with_pivot(
                df, out_path, sheet_name=args.sheet or "Inc", taxonomy_result=result,
            )
        else:
            save_results(df, out_path)
        if args.save_taxonomy:
            from textanalyzer.services.io import IOService
            IOService.save_taxonomy(result, args.save_taxonomy)
            print(f"Saved taxonomy to: {args.save_taxonomy}")
        stats = result.stats
        print(
            f"Discovered {stats['n_subclusters']} subcategories "
            f"({stats['pct_non_repetitive']:.0%} Non-Repetitive). Saved to: {out_path}"
        )
        return

    # Load-and-apply path: skip training entirely.
    if args.load_model:
        if not args.model_path:
            raise ValueError("--load_model requires --model_path pointing to a saved joblib file")
        if not os.path.isfile(args.model_path):
            raise FileNotFoundError(f"Model file not found: {args.model_path}")
        payload = joblib.load(args.model_path)
        loaded_model = payload.get("model")
        loaded_vectorizer = payload.get("vectorizer")
        loaded_cluster_names = payload.get("cluster_names", {})
        if loaded_model is None or loaded_vectorizer is None:
            raise RuntimeError("Loaded payload missing 'model' or 'vectorizer'.")
        X_new = loaded_vectorizer.transform(processed)
        if not hasattr(loaded_model, "predict"):
            raise RuntimeError(
                "Loaded model does not support .predict(). Only KMeans/MiniBatchKMeans models — "
                "or models wrapped with ApplicableModel — can be applied to new data."
            )
        labels = loaded_model.predict(X_new)
        df["cluster_label"] = labels
        df["cluster_name"] = [loaded_cluster_names.get(int(l), "") for l in labels]
        input_ext = get_file_extension(args.input)
        default_ext = input_ext if input_ext in {".xlsx", ".csv", ".json"} else ".xlsx"
        out_path = args.output or os.path.splitext(args.input)[0] + "_clustered_from_model" + default_ext
        save_results(df, out_path)
        print(f"Applied loaded model and saved results to: {out_path}")
        return

    # If all texts are empty after preprocessing, handle gracefully
    if not any(s.strip() for s in processed):
        print("Warning: all texts are empty after preprocessing. Creating a default label of -1 for all rows.")
        df["cluster_label"] = -1
        input_ext = get_file_extension(args.input)
        default_ext = input_ext if input_ext in {".xlsx", ".csv", ".json"} else ".xlsx"
        out_path = args.output or os.path.splitext(args.input)[0] + "_clustered" + default_ext
        save_results(df, out_path)
        return

    vectorizer, X = vectorize_texts(
        processed,
        max_features=args.max_features,
        use_hashing=args.use_hashing,
        chunk_size=args.chunk_size,
        vectorizer_kind=args.vectorizer,
        embedding_model=args.embedding_model,
        embedding_device=args.embedding_device,
        embedding_batch_size=args.embedding_batch_size,
    )

    model, labels = cluster_texts(
        X,
        algorithm=args.algorithm,
        n_clusters=args.n_clusters,
        eps=args.eps,
        min_samples=args.min_samples,
        min_cluster_size=args.min_cluster_size,
        use_gpu=args.use_gpu,
        minibatch_threshold=args.minibatch_threshold,
    )

    df["cluster_label"] = labels

    input_ext = get_file_extension(args.input)
    default_ext = input_ext if input_ext in {".xlsx", ".csv", ".json"} else ".xlsx"
    out_path = args.output or os.path.splitext(args.input)[0] + "_clustered" + default_ext
    save_results(df, out_path)

    # Top keywords per cluster and assign descriptive names
    cluster_names = {}
    try:
        top_keywords = get_top_keywords_per_cluster(vectorizer, X, labels, top_n=args.top_n)
        print("Top keywords per cluster:")
        for cluster_id, terms in top_keywords.items():
            print(f"Cluster {cluster_id}: ", ", ".join([t for t, s in terms]))

        # Assign human-readable cluster names
        cluster_names = assign_cluster_names(top_keywords, name_top_n=args.name_top_n, joiner=args.name_joiner)
        print("Assigned cluster names:")
        for cid, name in cluster_names.items():
            print(f"  {cid} -> {name}")

        # Map names into dataframe
        df["cluster_name"] = [cluster_names.get(int(l), "") for l in labels]
    except Exception as e:
        print(f"Could not compute top keywords or assign names: {e}")
        df["cluster_name"] = ""

    # Visualization
    if args.visualize:
        vis_out = os.path.splitext(out_path)[0] + f"_vis_{args.vis_method}.png"
        try:
            visualize_embeddings(X, labels, method=args.vis_method, out_path=vis_out)
        except Exception as e:
            print(f"Visualization failed: {e}")

    # Save model
    if args.save_model:
        model_path = args.model_path or os.path.splitext(out_path)[0] + "_model.joblib"
        try:
            # Wrap DBSCAN/Agglomerative/HDBSCAN so the saved artifact can still
            # .predict() on new data (via nearest-neighbor lookup over X).
            saved_model = model
            if args.algorithm in ("dbscan", "agglomerative", "hdbscan"):
                saved_model = wrap_model_for_prediction(model, X, labels, args.algorithm)
            joblib.dump({
                "model": saved_model,
                "vectorizer": vectorizer,
                "vectorizer_kind": args.vectorizer,
                "cluster_names": cluster_names,
                "top_keywords": top_keywords,
                "algorithm": args.algorithm,
            }, model_path)
            print(f"Saved model+vectorizer+names to: {model_path}")
        except Exception as e:
            print(f"Failed to save model: {e}")


if __name__ == "__main__":
    main()
