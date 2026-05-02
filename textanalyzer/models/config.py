"""Cleaning-config model — bridges UI checkbox state to ``TextCleaningConfig``."""

from __future__ import annotations

from dataclasses import asdict, dataclass

from cluster_tool import TextCleaningConfig


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


__all__ = ["CleaningConfigModel"]
