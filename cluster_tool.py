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
import os
import re
import warnings
from dataclasses import dataclass, field
from datetime import datetime
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
import matplotlib.pyplot as plt
import seaborn as sns
import joblib

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

# Optional NLTK lemmatization.
try:
    import nltk  # type: ignore[import-untyped]
    from nltk.stem import WordNetLemmatizer  # type: ignore[import-untyped]
    from nltk.tokenize import word_tokenize  # type: ignore[import-untyped]
    _NLTK_AVAILABLE = True
except ImportError:
    _NLTK_AVAILABLE = False

# Cap silhouette computations on huge datasets so metrics stay tractable.
_MAX_SILHOUETTE_SAMPLES = 5000

# One-time NLTK data download flag.
_NLTK_DATA_READY = False


URL_RE = re.compile(r"https?://\S+|www\.\S+", flags=re.IGNORECASE)
EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", flags=re.IGNORECASE)
PUNCT_RE = re.compile(r"[^\w\s]")
NUMBER_RE = re.compile(r"\d+")
WHITESPACE_RE = re.compile(r"\s+")


@dataclass
class TextCleaningConfig:
    replace_missing: bool = True
    missing_value_text: str = ""
    trim_whitespace: bool = True
    lowercase: bool = True
    collapse_whitespace: bool = True
    remove_punctuation: bool = False
    remove_numbers: bool = False
    remove_urls: bool = False
    remove_emails: bool = False
    regex_pattern: str = ""
    regex_replacement: str = ""
    dedupe_cleaned_rows: bool = False
    lemmatize: bool = False
    custom_stopwords: Tuple[str, ...] = ()

    def __post_init__(self):
        self.missing_value_text = str(self.missing_value_text)
        self.regex_pattern = str(self.regex_pattern or "")
        self.regex_replacement = str(self.regex_replacement or "")
        if self.regex_pattern:
            try:
                re.compile(self.regex_pattern)
            except re.error as exc:
                raise ValueError(f"Invalid regex pattern: {exc}") from exc


@dataclass
class TextCleaningResult:
    cleaned_texts: List[str]
    cluster_input_texts: List[str]
    kept_indices: List[int]
    representative_index_by_row: List[Optional[int]]
    stats: Dict[str, Any]
    preview_rows: List[Dict[str, str]] = field(default_factory=list)


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


def coerce_text_column(series: pd.Series) -> pd.Series:
    # Convert to string, preserving NaN as empty strings
    return series.fillna("").astype(str)


def get_default_text_cleaning_config() -> TextCleaningConfig:
    return TextCleaningConfig()


def preprocess_texts(texts: List[str], config: Optional[TextCleaningConfig] = None) -> List[str]:
    config = config or get_default_text_cleaning_config()
    return [clean_text_value(text, config) for text in texts]


def clean_text_value(value: Any, config: Optional[TextCleaningConfig] = None) -> str:
    config = config or get_default_text_cleaning_config()

    if pd.isna(value):
        text = config.missing_value_text if config.replace_missing else ""
    else:
        text = str(value)

    if config.trim_whitespace:
        text = text.strip()
    if config.lowercase:
        text = text.lower()
    if config.remove_urls:
        text = URL_RE.sub(" ", text)
    if config.remove_emails:
        text = EMAIL_RE.sub(" ", text)
    if config.regex_pattern:
        text = re.sub(config.regex_pattern, config.regex_replacement, text)
    if config.remove_punctuation:
        text = PUNCT_RE.sub(" ", text)
    if config.remove_numbers:
        text = NUMBER_RE.sub(" ", text)
    if config.collapse_whitespace:
        text = WHITESPACE_RE.sub(" ", text)
    if config.trim_whitespace or config.collapse_whitespace:
        text = text.strip()
    if config.lemmatize:
        text = _apply_lemmatization(text)
    if config.custom_stopwords and text:
        stops = {sw.lower() for sw in config.custom_stopwords}
        text = " ".join(tok for tok in text.split() if tok not in stops)
    return text


def _apply_lemmatization(text: str) -> str:
    """Lemmatize a text string using NLTK's WordNetLemmatizer.

    Lazily downloads required NLTK corpora the first time it's called. Returns
    the original text unchanged if NLTK is unavailable or the download fails —
    lemmatization is a best-effort enhancement, not a hard requirement.
    """
    if not text or not _NLTK_AVAILABLE:
        return text
    global _NLTK_DATA_READY
    if not _NLTK_DATA_READY:
        try:
            nltk.data.find("tokenizers/punkt")
            nltk.data.find("corpora/wordnet")
            _NLTK_DATA_READY = True
        except LookupError:
            try:
                nltk.download("punkt", quiet=True)
                nltk.download("punkt_tab", quiet=True)
                nltk.download("wordnet", quiet=True)
                _NLTK_DATA_READY = True
            except Exception:
                warnings.warn("Could not download NLTK data. Lemmatization disabled.")
                return text
    try:
        lemmatizer = WordNetLemmatizer()
        tokens = word_tokenize(text)
        return " ".join(lemmatizer.lemmatize(tok) for tok in tokens)
    except Exception:
        return text


def prepare_text_cleaning(texts: List[Any], config: Optional[TextCleaningConfig] = None, sample_size: int = 5) -> TextCleaningResult:
    config = config or get_default_text_cleaning_config()
    raw_texts = ["" if pd.isna(value) else str(value) for value in texts]
    cleaned_texts = [clean_text_value(value, config) for value in texts]

    kept_indices: List[int] = []
    representative_index_by_row: List[Optional[int]] = []
    seen_cleaned: Dict[str, int] = {}
    deduped_row_count = 0

    for index, cleaned in enumerate(cleaned_texts):
        if not cleaned:
            representative_index_by_row.append(None)
            continue

        if config.dedupe_cleaned_rows and cleaned in seen_cleaned:
            representative_index_by_row.append(seen_cleaned[cleaned])
            deduped_row_count += 1
            continue

        seen_cleaned[cleaned] = index
        kept_indices.append(index)
        representative_index_by_row.append(index)

    cluster_input_texts = [cleaned_texts[index] for index in kept_indices]
    empty_count = sum(1 for item in cleaned_texts if not item)
    preview_rows = [
        {"raw": raw_texts[index], "cleaned": cleaned_texts[index]}
        for index in range(min(sample_size, len(cleaned_texts)))
    ]

    stats = {
        "source_row_count": len(raw_texts),
        "cleaned_row_count": len(cleaned_texts),
        "kept_row_count": len(kept_indices),
        "deduped_row_count": deduped_row_count,
        "empty_row_count": empty_count,
        "cleaned_column_name_suffix": "_cleaned",
    }
    return TextCleaningResult(
        cleaned_texts=cleaned_texts,
        cluster_input_texts=cluster_input_texts,
        kept_indices=kept_indices,
        representative_index_by_row=representative_index_by_row,
        stats=stats,
        preview_rows=preview_rows,
    )


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
) -> Tuple[Any, Any]:
    """Vectorize texts into TF-IDF features.

    Default behavior is unchanged from the previous single-arg signature. Pass
    `use_hashing=True` for a stateless, memory-bounded path that streams
    documents in chunks (useful for very large corpora). When hashing is on the
    returned vectorizer is a `HashingVectorizer` and `get_top_keywords_per_cluster`
    will return empty keyword lists since hashing has no feature names.
    """
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
                X_dense = X.toarray() if hasattr(X, "toarray") else X
                X_dense = X_dense.astype("float32")
                model = cuml.cluster.KMeans(n_clusters=n_clusters, random_state=random_state)
                model.fit(X_dense)
                return model, model.predict(X_dense)
            except Exception:
                pass  # fall through to CPU path
        if n_samples >= minibatch_threshold:
            model = MiniBatchKMeans(n_clusters=n_clusters, random_state=random_state, n_init=10)
        else:
            model = KMeans(n_clusters=n_clusters, random_state=random_state, n_init=10)
        labels = model.fit_predict(X)
        return model, labels

    if algorithm == "dbscan":
        model = DBSCAN(eps=eps, min_samples=min_samples, metric="cosine")
        return model, model.fit_predict(X)

    if algorithm == "hdbscan":
        if not _HDBSCAN_AVAILABLE:
            raise ImportError("HDBSCAN not installed. Install with: pip install hdbscan")
        X_dense = X.toarray() if hasattr(X, "toarray") else X
        model = hdbscan.HDBSCAN(
            min_cluster_size=min_cluster_size,
            min_samples=min_samples,
            metric="euclidean",
            cluster_selection_method="eom",
        )
        return model, model.fit_predict(X_dense)

    if algorithm == "agglomerative":
        model = AgglomerativeClustering(n_clusters=n_clusters)
        return model, model.fit_predict(X.toarray() if hasattr(X, "toarray") else X)

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
            model = KMeans(n_clusters=k, random_state=random_state, n_init=10)
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
                model = KMeans(n_clusters=n_clusters, random_state=random_state, n_init=10)
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
                labels = model.fit_predict(X.toarray() if hasattr(X, "toarray") else X)

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
                    Xm_dense = Xm.toarray() if hasattr(Xm, "toarray") else Xm
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
            X_dense = X_train.toarray() if hasattr(X_train, "toarray") else X_train
            self._nn = NearestNeighbors(n_neighbors=1, metric=self._NN_METRIC)
            self._nn.fit(X_dense)
            self._train_labels = self.labels

    def predict(self, X):
        if self.algorithm in ("kmeans",) and hasattr(self.model, "predict"):
            return self.model.predict(X)
        if self._nn is not None and self._train_labels is not None:
            X_dense = X.toarray() if hasattr(X, "toarray") else X
            _, indices = self._nn.kneighbors(X_dense)
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
    """
    result: Dict[int, List[Tuple[str, float]]] = {}
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


def compute_embedding_2d(X, method: str = "pca", perplexity: int = 30, random_state: int = 42) -> np.ndarray:
    """Return a 2D embedding of X using PCA or t-SNE. Reusable by GUI for inline rendering."""
    dense = X.toarray() if hasattr(X, "toarray") else X
    if method == "pca":
        reducer = PCA(n_components=2, random_state=random_state)
    elif method == "tsne":
        effective_perplexity = max(2, min(perplexity, max(2, dense.shape[0] - 1)))
        reducer = TSNE(n_components=2, perplexity=effective_perplexity, random_state=random_state)
    else:
        raise ValueError("Unknown visualization method: choose 'pca' or 'tsne'")
    return reducer.fit_transform(dense)


def visualize_embeddings(X, labels: np.ndarray, method: str = "pca", perplexity: int = 30, random_state: int = 42, out_path: Optional[str] = None):
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

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        raise FileNotFoundError(f"Input file not found: {args.input}")

    df = load_table(args.input, sheet_name=args.sheet)
    if args.column not in df.columns:
        raise ValueError(f"Column '{args.column}' not found in input data. Available columns: {list(df.columns)}")

    text_series = coerce_text_column(df[args.column])
    processed = preprocess_texts(text_series.tolist())

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
                "cluster_names": cluster_names,
                "top_keywords": top_keywords,
                "algorithm": args.algorithm,
            }, model_path)
            print(f"Saved model+vectorizer+names to: {model_path}")
        except Exception as e:
            print(f"Failed to save model: {e}")


if __name__ == "__main__":
    main()
