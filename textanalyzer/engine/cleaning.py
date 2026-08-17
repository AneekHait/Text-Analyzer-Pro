"""Text cleaning pipeline — extracted from cluster.py for modularity.

All public symbols are re-exported by ``textanalyzer.engine.cluster`` so
existing imports continue to work unchanged.
"""

import re
import warnings
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


URL_RE = re.compile(r"https?://\S+|www\.\S+", flags=re.IGNORECASE)
EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", flags=re.IGNORECASE)
PUNCT_RE = re.compile(r"[^\w\s]")
NUMBER_RE = re.compile(r"\d+")
WHITESPACE_RE = re.compile(r"\s+")

_NLTK_DATA_READY = False
_LEMMATIZER_INSTANCE = None


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


def get_default_text_cleaning_config() -> TextCleaningConfig:
    return TextCleaningConfig()


def coerce_text_column(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str)


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


def _get_lemmatizer():
    """Return a cached WordNetLemmatizer, creating it on first call."""
    global _LEMMATIZER_INSTANCE, _NLTK_DATA_READY
    if _LEMMATIZER_INSTANCE is not None:
        return _LEMMATIZER_INSTANCE
    try:
        import nltk
        from nltk.stem import WordNetLemmatizer
    except ImportError:
        return None

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
                return None

    _LEMMATIZER_INSTANCE = WordNetLemmatizer()
    return _LEMMATIZER_INSTANCE


def _apply_lemmatization(text: str) -> str:
    """Lemmatize using a cached WordNetLemmatizer instance."""
    if not text:
        return text
    lemmatizer = _get_lemmatizer()
    if lemmatizer is None:
        return text
    try:
        from nltk.tokenize import word_tokenize
        tokens = word_tokenize(text)
        return " ".join(lemmatizer.lemmatize(tok) for tok in tokens)
    except Exception:
        return text


def prepare_text_cleaning(
    texts: List[Any], config: Optional[TextCleaningConfig] = None, sample_size: int = 5
) -> TextCleaningResult:
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
