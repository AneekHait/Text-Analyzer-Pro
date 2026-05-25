"""Cleaning-config model — bridges UI checkbox state to ``TextCleaningConfig``."""

from __future__ import annotations

from dataclasses import asdict, dataclass

from textanalyzer.engine.cluster import DEFAULT_EMBEDDING_MODEL, TextCleaningConfig


@dataclass
class CleaningConfigModel:
    """Mirror of :class:`TextCleaningConfig` that can round-trip to/from dict.

    The class exists so that GUI code never directly imports *cluster_tool*;
    the services layer converts between the two representations.
    """

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

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------
    def to_engine_config(self) -> TextCleaningConfig:
        """Return the *cluster_tool* ``TextCleaningConfig`` equivalent."""
        return TextCleaningConfig(**asdict(self))

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "CleaningConfigModel":
        known = {f.name for f in cls.__dataclass_fields__.values()}
        return cls(**{k: v for k, v in d.items() if k in known})


@dataclass
class CategorizationConfig:
    """User-facing knobs for the Run Categorization workflow.

    Mirrors the dialog fields. ``granularity`` is the primary control (0..100,
    Coarse → Fine); ``min_cluster_size`` is the underlying knob the slider
    drives. Both are persisted so the dialog opens with the user's last choice.
    """

    granularity: int = 50
    min_cluster_size: int = 5
    min_samples: int = 3
    non_repetitive_min_size: int = 5
    vectorizer_kind: str = "embedding"  # "embedding" | "tfidf"
    name_ngram_low: int = 1
    name_ngram_high: int = 3
    confidence_threshold: float = 0.45  # used by apply_taxonomy

    @staticmethod
    def min_cluster_size_from_granularity(g: int) -> int:
        """Map slider value 0..100 to HDBSCAN min_cluster_size.

        High granularity → smaller (more, finer) clusters. Mirrors the formula
        documented in the dialog mockup so tests can assert the mapping.
        """
        g = max(0, min(100, int(g)))
        return max(3, round(20 * (1 - g / 100) + 4))

    def to_engine_kwargs(self) -> dict:
        return {
            "vectorizer_kind": self.vectorizer_kind,
            "min_cluster_size": int(self.min_cluster_size),
            "min_samples": int(self.min_samples),
            "non_repetitive_min_size": int(self.non_repetitive_min_size),
            "name_ngram_range": (int(self.name_ngram_low), int(self.name_ngram_high)),
        }

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "CategorizationConfig":
        known = {f.name for f in cls.__dataclass_fields__.values()}
        return cls(**{k: v for k, v in d.items() if k in known})


@dataclass
class VectorizerConfig:
    """User-facing vectorization config — bridges UI state to ``vectorize_texts`` kwargs.

    ``kind`` selects between TF-IDF (default, lexical) and sentence-transformer
    embeddings (semantic). The embedding fields are only consulted when
    ``kind == "embedding"``.
    """

    kind: str = "tfidf"
    embedding_model: str = DEFAULT_EMBEDDING_MODEL
    embedding_device: str = "cpu"
    embedding_batch_size: int = 32

    def to_vectorize_kwargs(self) -> dict:
        """Render the subset of kwargs ``vectorize_texts`` understands."""
        return {
            "vectorizer_kind": self.kind,
            "embedding_model": self.embedding_model,
            "embedding_device": self.embedding_device,
            "embedding_batch_size": self.embedding_batch_size,
        }

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "VectorizerConfig":
        known = {f.name for f in cls.__dataclass_fields__.values()}
        return cls(**{k: v for k, v in d.items() if k in known})


__all__ = ["CategorizationConfig", "CleaningConfigModel", "VectorizerConfig"]
