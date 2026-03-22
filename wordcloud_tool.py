#!/usr/bin/env python3
"""
Utilities for building and exporting word clouds from a selected text column.
"""

from __future__ import annotations

import os
import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Sequence, Set, Tuple

import pandas as pd
from sklearn.feature_extraction.text import ENGLISH_STOP_WORDS

from cluster_tool import coerce_text_column

try:
    from wordcloud import WordCloud
except ImportError:  # pragma: no cover - exercised in runtime environments without the optional dependency
    WordCloud = None


TOKEN_RE = re.compile(r"\b[\w']+\b", flags=re.UNICODE)
PHRASE_MODE_TO_NGRAM = {
    "Unigrams": 1,
    "Up to Bigrams": 2,
    "Up to Trigrams": 3,
    "unigrams": 1,
    "bigrams": 2,
    "trigrams": 3,
}


@dataclass
class WordCloudConfig:
    max_words: int = 200
    min_frequency: int = 1
    width: int = 1200
    height: int = 700
    phrase_mode: str = "Unigrams"
    use_builtin_stopwords: bool = True
    lowercase: bool = True
    exclude_numeric: bool = True
    background_color: str = "white"
    colormap: str = "viridis"
    custom_stopwords: Set[str] = field(default_factory=set)

    def __post_init__(self):
        for attr_name in ("max_words", "min_frequency", "width", "height"):
            value = int(getattr(self, attr_name))
            if value < 1:
                raise ValueError(f"{attr_name} must be at least 1")
            setattr(self, attr_name, value)
        if self.phrase_mode not in PHRASE_MODE_TO_NGRAM:
            raise ValueError(f"Unsupported phrase mode: {self.phrase_mode}")
        self.custom_stopwords = {
            _normalize_stopword(word) for word in self.custom_stopwords if _normalize_stopword(word)
        }


def get_effective_stopwords(config: WordCloudConfig) -> Set[str]:
    stopwords: Set[str] = set()
    if config.use_builtin_stopwords:
        stopwords.update(ENGLISH_STOP_WORDS)
    stopwords.update(config.custom_stopwords)
    return stopwords


def build_term_stats(texts: Sequence[str], config: WordCloudConfig) -> pd.DataFrame:
    stats_df, _summary = prepare_wordcloud_data(texts, config)
    return stats_df


def summarize_texts(texts: Sequence[str], config: WordCloudConfig) -> Dict[str, int]:
    _stats_df, summary = prepare_wordcloud_data(texts, config)
    return summary


def prepare_wordcloud_data(texts: Sequence[str], config: WordCloudConfig) -> Tuple[pd.DataFrame, Dict[str, int]]:
    counts, summary = _collect_term_counts(texts, config)
    stats_df = _counts_to_dataframe(counts, config.min_frequency)
    summary.update(
        {
            "unique_terms": int(len(stats_df)),
            "kept_term_occurrences": int(stats_df["count"].sum()) if not stats_df.empty else 0,
        }
    )
    return stats_df, summary


def render_wordcloud(stats_df: pd.DataFrame, config: WordCloudConfig):
    if WordCloud is None:
        raise ImportError(
            "The 'wordcloud' package is required for preview rendering. Install it with 'pip install -r requirements.txt'."
        )
    if stats_df.empty:
        raise ValueError("No terms are available to render after applying the current filters.")

    frequencies = dict(zip(stats_df["term"], stats_df["count"]))
    generator = WordCloud(
        width=config.width,
        height=config.height,
        background_color=config.background_color,
        colormap=config.colormap,
        max_words=config.max_words,
        collocations=False,
    )
    generator.generate_from_frequencies(frequencies)
    return generator.to_image()


def export_term_stats(stats_df: pd.DataFrame, out_path: str) -> str:
    try:
        stats_df.to_excel(out_path, index=False, engine="openpyxl")
        return out_path
    except PermissionError:
        base, ext = os.path.splitext(out_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt_path = f"{base}_writable_{timestamp}{ext}"
        stats_df.to_excel(alt_path, index=False, engine="openpyxl")
        return alt_path


def _collect_term_counts(texts: Sequence[str], config: WordCloudConfig) -> Tuple[Counter, Dict[str, int]]:
    series = coerce_text_column(pd.Series(list(texts)))
    stopwords = get_effective_stopwords(config)
    max_ngram = PHRASE_MODE_TO_NGRAM[config.phrase_mode]

    counts: Counter = Counter()
    total_rows = int(len(series))
    usable_rows = 0
    total_term_occurrences = 0

    for text in series.tolist():
        tokens = _tokenize_text(text, config, stopwords)
        if not tokens:
            continue

        row_terms = _build_row_terms(tokens, max_ngram)
        if not row_terms:
            continue

        usable_rows += 1
        total_term_occurrences += len(row_terms)
        counts.update(row_terms)

    return counts, {
        "total_rows": total_rows,
        "usable_rows": usable_rows,
        "term_occurrences": total_term_occurrences,
    }


def _tokenize_text(text: str, config: WordCloudConfig, stopwords: Set[str]) -> List[str]:
    tokens: List[str] = []
    for raw_token in TOKEN_RE.findall(text):
        display_token = raw_token.lower() if config.lowercase else raw_token
        normalized_token = raw_token.lower()

        if config.exclude_numeric and normalized_token.isdigit():
            continue
        if normalized_token in stopwords:
            continue

        tokens.append(display_token)
    return tokens


def _build_row_terms(tokens: Sequence[str], max_ngram: int) -> List[str]:
    row_terms: List[str] = []
    for ngram_size in range(1, max_ngram + 1):
        if len(tokens) < ngram_size:
            continue
        for index in range(len(tokens) - ngram_size + 1):
            row_terms.append(" ".join(tokens[index:index + ngram_size]))
    return row_terms


def _counts_to_dataframe(counts: Counter, min_frequency: int) -> pd.DataFrame:
    filtered_items = [
        (term, count)
        for term, count in counts.items()
        if count >= min_frequency
    ]
    filtered_items.sort(key=lambda item: (-item[1], item[0]))

    if not filtered_items:
        return pd.DataFrame(columns=["term", "count", "share"])

    kept_total = sum(count for _term, count in filtered_items)
    rows = [
        {
            "term": term,
            "count": int(count),
            "share": count / kept_total,
        }
        for term, count in filtered_items
    ]
    return pd.DataFrame(rows, columns=["term", "count", "share"])


def _normalize_stopword(word: str) -> str:
    return str(word).strip().lower()
