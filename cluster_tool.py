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
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans, DBSCAN, AgglomerativeClustering
from sklearn.decomposition import PCA
from sklearn.manifold import TSNE
import matplotlib.pyplot as plt
import seaborn as sns
import joblib


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


SUPPORTED_INPUT_EXTENSIONS = {".xlsx", ".xls", ".csv", ".json"}
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


def load_table(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    ext = get_file_extension(path)
    if ext in {".xlsx", ".xls"}:
        if sheet_name is None:
            return pd.read_excel(path, engine="openpyxl")
        return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    if ext == ".csv":
        return pd.read_csv(path)
    if ext == ".json":
        try:
            return _coerce_json_frame(pd.read_json(path))
        except ValueError:
            return _coerce_json_frame(pd.read_json(path, lines=True))
    raise ValueError(
        f"Unsupported file type '{ext or '(none)'}'. Supported types: Excel (.xlsx, .xls), CSV (.csv), JSON (.json)."
    )


def get_sheet_names(path: str) -> List[str]:
    ext = get_file_extension(path)
    if ext in {".xlsx", ".xls"}:
        workbook = pd.ExcelFile(path, engine="openpyxl")
        return workbook.sheet_names
    if ext in {".csv", ".json"}:
        return [SINGLE_TABLE_SHEET_NAME]
    raise ValueError(
        f"Unsupported file type '{ext or '(none)'}'. Supported types: Excel (.xlsx, .xls), CSV (.csv), JSON (.json)."
    )


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


def vectorize_texts(texts: List[str], max_features: Optional[int] = 2000) -> Tuple[TfidfVectorizer, np.ndarray]:
    vectorizer = TfidfVectorizer(stop_words="english", max_features=max_features)
    X = vectorizer.fit_transform(texts)
    return vectorizer, X


def cluster_texts(
    X, algorithm: str = "kmeans", n_clusters: int = 5, eps: float = 0.5, min_samples: int = 5, random_state: int = 42
):
    if algorithm == "kmeans":
        model = KMeans(n_clusters=n_clusters, random_state=random_state)
        labels = model.fit_predict(X)
    elif algorithm == "dbscan":
        # DBSCAN requires dense or sparse; works with sparse
        model = DBSCAN(eps=eps, min_samples=min_samples, metric="cosine")
        labels = model.fit_predict(X)
    elif algorithm == "agglomerative":
        # AgglomerativeClustering does not accept sparse matrices -> densify
        model = AgglomerativeClustering(n_clusters=n_clusters)
        labels = model.fit_predict(X.toarray())
    else:
        raise ValueError(f"Unknown algorithm: {algorithm}")
    return model, labels


def get_top_keywords_per_cluster(
    vectorizer: TfidfVectorizer, X, labels: np.ndarray, top_n: int = 10
) -> Dict[int, List[Tuple[str, float]]]:
    features = vectorizer.get_feature_names_out()
    df_tfidf = pd.DataFrame(X.todense(), columns=features)
    result = {}
    for label in sorted(np.unique(labels)):
        if label == -1:
            # noise for DBSCAN
            continue
        mask = labels == label
        if mask.sum() == 0:
            result[label] = []
            continue
        mean_tfidf = df_tfidf[mask].mean(axis=0)
        top_idx = mean_tfidf.nlargest(top_n).index
        top_scores = mean_tfidf.loc[top_idx].values
        result[int(label)] = list(zip(top_idx.tolist(), top_scores.tolist()))
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


def visualize_embeddings(X, labels: np.ndarray, method: str = "pca", perplexity: int = 30, random_state: int = 42, out_path: Optional[str] = None):
    if method == "pca":
        reducer = PCA(n_components=2, random_state=random_state)
        emb = reducer.fit_transform(X.toarray() if hasattr(X, "toarray") else X)
    elif method == "tsne":
        reducer = TSNE(n_components=2, perplexity=perplexity, random_state=random_state)
        emb = reducer.fit_transform(X.toarray() if hasattr(X, "toarray") else X)
    else:
        raise ValueError("Unknown visualization method: choose 'pca' or 'tsne'")

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
    parser.add_argument("--algorithm", "-a", choices=["kmeans", "dbscan", "agglomerative"], default="kmeans")
    parser.add_argument("--n_clusters", "-k", type=int, default=5, help="Number of clusters (for kmeans/agglomerative)")
    parser.add_argument("--eps", type=float, default=0.5, help="DBSCAN eps parameter")
    parser.add_argument("--min_samples", type=int, default=5, help="DBSCAN min_samples")
    parser.add_argument("--max_features", type=int, default=2000, help="Max features for TF-IDF")
    parser.add_argument("--output", "-o", default=None, help="Output path (optional) — defaults to input + _clustered with a matching supported extension")
    parser.add_argument("--visualize", "-v", action="store_true", help="Visualize clusters (PCA)")
    parser.add_argument("--vis_method", choices=["pca", "tsne"], default="pca", help="Visualization method")
    parser.add_argument("--top_n", type=int, default=10, help="Top keywords per cluster")
    parser.add_argument("--name_top_n", type=int, default=3, help="Number of top keywords to form cluster name")
    parser.add_argument("--name_joiner", type=str, default=", ", help="String to join keywords when forming cluster name")
    parser.add_argument("--save_model", action="store_true", help="Save trained clustering model (joblib)")
    parser.add_argument("--model_path", default=None, help="Path to save the model (optional)")

    args = parser.parse_args()

    if not os.path.isfile(args.input):
        raise FileNotFoundError(f"Input file not found: {args.input}")

    df = load_table(args.input, sheet_name=args.sheet)
    if args.column not in df.columns:
        raise ValueError(f"Column '{args.column}' not found in input data. Available columns: {list(df.columns)}")

    text_series = coerce_text_column(df[args.column])
    processed = preprocess_texts(text_series.tolist())

    # (No repeated-keyword extraction in this version)

    # If all texts are empty after preprocessing, handle gracefully
    if not any(s.strip() for s in processed):
        print("Warning: all texts are empty after preprocessing. Creating a default label of -1 for all rows.")
        df["cluster_label"] = -1
        input_ext = get_file_extension(args.input)
        default_ext = input_ext if input_ext in {".xlsx", ".csv", ".json"} else ".xlsx"
        out_path = args.output or os.path.splitext(args.input)[0] + "_clustered" + default_ext
        save_results(df, out_path)
        return

    vectorizer, X = vectorize_texts(processed, max_features=args.max_features)

    model, labels = cluster_texts(
        X, algorithm=args.algorithm, n_clusters=args.n_clusters, eps=args.eps, min_samples=args.min_samples
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
            # Save model, vectorizer, and cluster name/keymapping for reuse
            joblib.dump({"model": model, "vectorizer": vectorizer, "cluster_names": cluster_names, "top_keywords": top_keywords}, model_path)
            print(f"Saved model+vectorizer+names to: {model_path}")
        except Exception as e:
            print(f"Failed to save model: {e}")


if __name__ == "__main__":
    main()
