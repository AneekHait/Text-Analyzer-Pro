#!/usr/bin/env python3
"""
Utilities for building, styling, and exporting word clouds from a selected text column.
"""

from __future__ import annotations

import json
import math
import os
import random
import re
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple

import numpy as np
import pandas as pd
from PIL import Image, ImageColor, ImageDraw
from sklearn.feature_extraction.text import ENGLISH_STOP_WORDS

from .cluster import coerce_text_column

try:
    from wordcloud import WordCloud
except ImportError:  # pragma: no cover - exercised in runtime environments without the optional dependency
    WordCloud = None


TOKEN_RE = re.compile(r"\b[\w']+\b", flags=re.UNICODE)
HEX_COLOR_RE = re.compile(r"^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")
PHRASE_MODE_TO_NGRAM = {
    "Unigrams": 1,
    "Up to Bigrams": 2,
    "Up to Trigrams": 3,
    "unigrams": 1,
    "bigrams": 2,
    "trigrams": 3,
}
COLOR_MODES = ("Colormap", "Palette", "Custom")
MASK_MODES = ("None", "Builtin Shape", "Custom PNG")
BUILTIN_SHAPES = (
    "Rectangle",
    "Circle",
    "Heart",
    "Star",
    "Speech Bubble",
    "Diamond",
    "Hexagon",
    "Triangle",
    "Shield",
    "Cloud",
)
PALETTES = {
    "Viridis": ["#440154", "#31688e", "#35b779", "#fde725"],
    "Monochrome": ["#111111", "#444444", "#777777", "#bdbdbd"],
    "Warm": ["#7f2704", "#d94801", "#f16913", "#fd8d3c", "#fdd0a2"],
    "Cool": ["#084081", "#0868ac", "#2b8cbe", "#4eb3d3", "#a8ddb5"],
    "Pastel": ["#ffcad4", "#f4acb7", "#9d8189", "#84a59d", "#f6bd60"],
    "High Contrast": ["#111111", "#e63946", "#ffb703", "#219ebc", "#f1faee"],
}
DEFAULT_PRESET_NAME = "Default"
SORT_MODES = ("Frequency", "Alphabetical")


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
    font_path: str = ""
    font_label: str = "Default font"
    color_mode: str = "Colormap"
    palette_name: str = "Viridis"
    custom_colors: List[str] = field(default_factory=list)
    contour_color: str = "#111111"
    contour_width: int = 0
    mask_mode: str = "None"
    shape_name: str = "Rectangle"
    mask_path: str = ""
    prefer_horizontal: float = 0.9
    repeat: bool = False
    relative_scaling: float = 0.5
    scale: int = 1
    custom_stopwords: Set[str] = field(default_factory=set)
    include_terms: Set[str] = field(default_factory=set)
    exclude_terms: Set[str] = field(default_factory=set)
    render_top_n: int = 0
    sort_mode: str = "Frequency"
    template_name: str = DEFAULT_PRESET_NAME

    def __post_init__(self):
        for attr_name in ("max_words", "min_frequency", "width", "height", "contour_width", "scale", "render_top_n"):
            value = int(getattr(self, attr_name))
            if value < 1 and attr_name not in {"contour_width", "render_top_n"}:
                raise ValueError(f"{attr_name} must be at least 1")
            if attr_name in {"contour_width", "render_top_n"} and value < 0:
                raise ValueError(f"{attr_name} must be at least 0")
            setattr(self, attr_name, value)

        for attr_name, minimum, maximum in (
            ("prefer_horizontal", 0.0, 1.0),
            ("relative_scaling", 0.0, 1.0),
        ):
            value = float(getattr(self, attr_name))
            if not minimum <= value <= maximum:
                raise ValueError(f"{attr_name} must be between {minimum} and {maximum}")
            setattr(self, attr_name, value)

        if self.phrase_mode not in PHRASE_MODE_TO_NGRAM:
            raise ValueError(f"Unsupported phrase mode: {self.phrase_mode}")
        if self.color_mode not in COLOR_MODES:
            raise ValueError(f"Unsupported color mode: {self.color_mode}")
        if self.mask_mode not in MASK_MODES:
            raise ValueError(f"Unsupported mask mode: {self.mask_mode}")
        if self.shape_name not in BUILTIN_SHAPES:
            raise ValueError(f"Unsupported shape name: {self.shape_name}")
        if self.palette_name not in PALETTES:
            raise ValueError(f"Unsupported palette: {self.palette_name}")
        if self.sort_mode not in SORT_MODES:
            raise ValueError(f"Unsupported sort mode: {self.sort_mode}")

        self.font_path = str(self.font_path or "").strip()
        self.mask_path = str(self.mask_path or "").strip()
        self.font_label = str(self.font_label or "").strip() or "Default font"
        self.template_name = str(self.template_name or DEFAULT_PRESET_NAME).strip() or DEFAULT_PRESET_NAME
        self.background_color = _normalize_color_value(self.background_color)
        self.contour_color = _normalize_color_value(self.contour_color)
        self.custom_colors = [_normalize_hex_color(color) for color in self.custom_colors if str(color).strip()]
        self.custom_stopwords = {
            _normalize_stopword(word) for word in self.custom_stopwords if _normalize_stopword(word)
        }
        self.include_terms = {
            _normalize_stopword(word) for word in self.include_terms if _normalize_stopword(word)
        }
        self.exclude_terms = {
            _normalize_stopword(word) for word in self.exclude_terms if _normalize_stopword(word)
        }

        if self.font_path and not os.path.isfile(self.font_path):
            raise ValueError(f"Font file does not exist: {self.font_path}")
        if self.mask_mode == "Custom PNG":
            if not self.mask_path:
                raise ValueError("Select a PNG mask file when mask mode is 'Custom PNG'.")
            if not os.path.isfile(self.mask_path):
                raise ValueError(f"Mask PNG does not exist: {self.mask_path}")
            if not self.mask_path.lower().endswith(".png"):
                raise ValueError("Custom masks must be PNG files.")
        if self.color_mode == "Custom" and not self.custom_colors:
            raise ValueError("Provide at least one custom hex color when color mode is 'Custom'.")


DEFAULT_TEMPLATE_PAYLOAD: Dict[str, object] = {
    "template_name": "Default",
    "background_color": "white",
    "color_mode": "Colormap",
    "colormap": "viridis",
    "palette_name": "Viridis",
    "mask_mode": "None",
    "shape_name": "Rectangle",
    "contour_width": 0,
    "prefer_horizontal": 0.9,
    "relative_scaling": 0.5,
    "scale": 1,
}


BUILTIN_TEMPLATES: Dict[str, Dict[str, object]] = {
    "Default": dict(DEFAULT_TEMPLATE_PAYLOAD),
    "Executive Clean": {
        **DEFAULT_TEMPLATE_PAYLOAD,
        "template_name": "Executive Clean",
        "background_color": "white",
        "color_mode": "Palette",
        "palette_name": "Monochrome",
        "contour_color": "#111111",
        "contour_width": 1,
        "mask_mode": "Builtin Shape",
        "shape_name": "Rectangle",
        "prefer_horizontal": 0.95,
    },
    "High Contrast": {
        **DEFAULT_TEMPLATE_PAYLOAD,
        "template_name": "High Contrast",
        "background_color": "black",
        "color_mode": "Palette",
        "palette_name": "High Contrast",
        "contour_color": "#f1faee",
        "contour_width": 2,
        "mask_mode": "Builtin Shape",
        "shape_name": "Circle",
        "prefer_horizontal": 0.85,
        "scale": 2,
    },
    "Poster": {
        **DEFAULT_TEMPLATE_PAYLOAD,
        "template_name": "Poster",
        "background_color": "ivory",
        "color_mode": "Palette",
        "palette_name": "Warm",
        "mask_mode": "Builtin Shape",
        "shape_name": "Star",
        "contour_color": "#7f2704",
        "contour_width": 2,
        "relative_scaling": 0.65,
        "scale": 2,
    },
    "Soft Editorial": {
        **DEFAULT_TEMPLATE_PAYLOAD,
        "template_name": "Soft Editorial",
        "background_color": "mintcream",
        "color_mode": "Palette",
        "palette_name": "Pastel",
        "mask_mode": "Builtin Shape",
        "shape_name": "Speech Bubble",
        "contour_color": "#84a59d",
        "contour_width": 1,
    },
    "Compact Dashboard": {
        **DEFAULT_TEMPLATE_PAYLOAD,
        "template_name": "Compact Dashboard",
        "background_color": "whitesmoke",
        "color_mode": "Palette",
        "palette_name": "Cool",
        "mask_mode": "Builtin Shape",
        "shape_name": "Circle",
        "prefer_horizontal": 1.0,
        "relative_scaling": 0.35,
        "scale": 1,
    },
}


def get_effective_stopwords(config: WordCloudConfig) -> Set[str]:
    stopwords: Set[str] = set()
    if config.use_builtin_stopwords:
        stopwords.update(ENGLISH_STOP_WORDS)
    stopwords.update(config.custom_stopwords)
    return stopwords


def get_palette_names() -> Tuple[str, ...]:
    return tuple(PALETTES.keys())


def get_builtin_shape_names() -> Tuple[str, ...]:
    return BUILTIN_SHAPES


def get_color_modes() -> Tuple[str, ...]:
    return COLOR_MODES


def get_mask_modes() -> Tuple[str, ...]:
    return MASK_MODES


def get_sort_modes() -> Tuple[str, ...]:
    return SORT_MODES


def get_template_names() -> Tuple[str, ...]:
    return tuple(BUILTIN_TEMPLATES.keys())


_FONT_CHOICES_CACHE: Tuple[Tuple[str, str], ...] = ()


def get_font_choices() -> Tuple[Tuple[str, str], ...]:
    global _FONT_CHOICES_CACHE
    if _FONT_CHOICES_CACHE:
        return _FONT_CHOICES_CACHE
    search_roots = []
    windir = os.environ.get("WINDIR")
    if windir:
        search_roots.append(os.path.join(windir, "Fonts"))
    search_roots.extend(
        [
            "/System/Library/Fonts",
            "/Library/Fonts",
            "/usr/share/fonts/truetype",
            "/usr/share/fonts",
        ]
    )

    candidates = [
        ("Default font", ""),
        ("Segoe UI", "segoeui.ttf"),
        ("Arial", "arial.ttf"),
        ("Calibri", "calibri.ttf"),
        ("Cambria", "cambria.ttc"),
        ("Georgia", "georgia.ttf"),
        ("Verdana", "verdana.ttf"),
        ("Trebuchet MS", "trebuc.ttf"),
        ("Times New Roman", "times.ttf"),
        ("Consolas", "consola.ttf"),
        ("Courier New", "cour.ttf"),
        ("Tahoma", "tahoma.ttf"),
        ("Impact", "impact.ttf"),
        ("Garamond", "gara.ttf"),
        ("Palatino", "pala.ttf"),
        ("Book Antiqua", "BOOKOS.TTF"),
        ("DejaVu Sans", "DejaVuSans.ttf"),
        ("DejaVu Serif", "DejaVuSerif.ttf"),
        ("Liberation Sans", "LiberationSans-Regular.ttf"),
        ("Liberation Serif", "LiberationSerif-Regular.ttf"),
        ("Noto Sans", "NotoSans-Regular.ttf"),
        ("Noto Serif", "NotoSerif-Regular.ttf"),
    ]

    resolved: List[Tuple[str, str]] = [("Default font", "")]
    seen_labels = {"Default font"}
    for label, filename in candidates[1:]:
        for root in search_roots:
            if not root or not os.path.isdir(root):
                continue
            for current_root, _dirs, files in os.walk(root):
                lowered = {file_name.lower(): file_name for file_name in files}
                lookup = lowered.get(filename.lower())
                if lookup:
                    if label not in seen_labels:
                        resolved.append((label, os.path.join(current_root, lookup)))
                        seen_labels.add(label)
                    break
            if label in seen_labels:
                break
    _FONT_CHOICES_CACHE = tuple(resolved)
    return _FONT_CHOICES_CACHE


def get_template_config(name: str) -> WordCloudConfig:
    clean_name = str(name).strip() or DEFAULT_PRESET_NAME
    payload = BUILTIN_TEMPLATES.get(clean_name)
    if payload is None:
        raise ValueError(f"Unknown template: {clean_name}")
    return WordCloudConfig(**payload)


def build_term_stats(texts: Sequence[str], config: WordCloudConfig) -> pd.DataFrame:
    stats_df, _summary = prepare_wordcloud_data(texts, config)
    return stats_df


def summarize_texts(texts: Sequence[str], config: WordCloudConfig) -> Dict[str, int]:
    _stats_df, summary = prepare_wordcloud_data(texts, config)
    return summary


def prepare_wordcloud_data(texts: Sequence[str], config: WordCloudConfig) -> Tuple[pd.DataFrame, Dict[str, int]]:
    counts, summary = _collect_term_counts(texts, config)
    stats_df = _counts_to_dataframe(counts, config.min_frequency)
    stats_df = _apply_term_filters(stats_df, config)
    summary.update(
        {
            "unique_terms": int(len(stats_df)),
            "kept_term_occurrences": int(stats_df["count"].sum()) if not stats_df.empty else 0,
        }
    )
    return stats_df, summary


def resolve_mask(config: WordCloudConfig) -> Optional[np.ndarray]:
    if config.mask_mode == "None":
        return None
    if config.mask_mode == "Builtin Shape":
        return build_builtin_mask(config.shape_name, config.width, config.height)
    return load_mask_from_png(config.mask_path, config.width, config.height)


def resolve_color_sequence(config: WordCloudConfig) -> List[str]:
    if config.color_mode == "Custom":
        return list(config.custom_colors)
    if config.color_mode == "Palette":
        return list(PALETTES[config.palette_name])
    return []


def build_color_func(config: WordCloudConfig):
    colors = resolve_color_sequence(config)
    if not colors:
        return None

    def color_func(word, font_size, position, orientation, random_state=None, **kwargs):
        seed = hash((word, font_size, position, orientation))
        chooser = random.Random(seed)
        return chooser.choice(colors)

    return color_func


def render_wordcloud(stats_df: pd.DataFrame, config: WordCloudConfig):
    if WordCloud is None:
        raise ImportError(
            "The 'wordcloud' package is required for preview rendering. Install it with 'pip install -r requirements.txt'."
        )
    if stats_df.empty:
        raise ValueError("No terms are available to render after applying the current filters.")

    frequencies = dict(zip(stats_df["term"], stats_df["count"]))
    mask = resolve_mask(config)
    generator = WordCloud(
        width=config.width,
        height=config.height,
        background_color=config.background_color,
        colormap=config.colormap if config.color_mode == "Colormap" else None,
        color_func=build_color_func(config),
        font_path=config.font_path or None,
        contour_color=config.contour_color,
        contour_width=config.contour_width,
        mask=mask,
        max_words=config.max_words,
        prefer_horizontal=config.prefer_horizontal,
        repeat=config.repeat,
        relative_scaling=config.relative_scaling,
        scale=config.scale,
        collocations=False,
    )
    generator.generate_from_frequencies(frequencies)
    return generator.to_image()


def build_builtin_mask(shape_name: str, width: int, height: int) -> np.ndarray:
    if shape_name not in BUILTIN_SHAPES:
        raise ValueError(f"Unsupported shape name: {shape_name}")

    image = Image.new("L", (width, height), 255)
    draw = ImageDraw.Draw(image)
    padding = max(12, min(width, height) // 18)

    if shape_name == "Rectangle":
        draw.rectangle((padding, padding, width - padding, height - padding), fill=0)
    elif shape_name == "Circle":
        draw.ellipse((padding, padding, width - padding, height - padding), fill=0)
    elif shape_name == "Heart":
        left_box = (width * 0.16, height * 0.12, width * 0.50, height * 0.46)
        right_box = (width * 0.50, height * 0.12, width * 0.84, height * 0.46)
        draw.ellipse(left_box, fill=0)
        draw.ellipse(right_box, fill=0)
        draw.polygon(
            [
                (width * 0.10, height * 0.32),
                (width * 0.50, height * 0.92),
                (width * 0.90, height * 0.32),
                (width * 0.74, height * 0.22),
                (width * 0.50, height * 0.46),
                (width * 0.26, height * 0.22),
            ],
            fill=0,
        )
    elif shape_name == "Star":
        center_x = width / 2
        center_y = height / 2
        outer_radius = min(width, height) * 0.42
        inner_radius = outer_radius * 0.42
        points = []
        for index in range(10):
            angle = math.radians(-90 + index * 36)
            radius = outer_radius if index % 2 == 0 else inner_radius
            points.append((center_x + radius * math.cos(angle), center_y + radius * math.sin(angle)))
        draw.polygon(points, fill=0)
    elif shape_name == "Speech Bubble":
        bubble_box = (padding, padding, width - padding, height * 0.74)
        draw.rounded_rectangle(bubble_box, radius=max(16, padding), fill=0)
        draw.polygon(
            [
                (width * 0.36, height * 0.74),
                (width * 0.48, height * 0.74),
                (width * 0.32, height * 0.92),
            ],
            fill=0,
        )
    elif shape_name == "Diamond":
        draw.polygon(
            [
                (width / 2, padding),
                (width - padding, height / 2),
                (width / 2, height - padding),
                (padding, height / 2),
            ],
            fill=0,
        )
    elif shape_name == "Hexagon":
        draw.polygon(
            [
                (width * 0.25, padding),
                (width * 0.75, padding),
                (width - padding, height / 2),
                (width * 0.75, height - padding),
                (width * 0.25, height - padding),
                (padding, height / 2),
            ],
            fill=0,
        )
    elif shape_name == "Triangle":
        draw.polygon(
            [
                (width / 2, padding),
                (width - padding, height - padding),
                (padding, height - padding),
            ],
            fill=0,
        )
    elif shape_name == "Shield":
        draw.polygon(
            [
                (width * 0.22, padding),
                (width * 0.78, padding),
                (width - padding, height * 0.30),
                (width * 0.82, height * 0.72),
                (width / 2, height - padding),
                (width * 0.18, height * 0.72),
                (padding, height * 0.30),
            ],
            fill=0,
        )
    elif shape_name == "Cloud":
        circles = [
            (width * 0.10, height * 0.35, width * 0.42, height * 0.72),
            (width * 0.28, height * 0.18, width * 0.60, height * 0.62),
            (width * 0.48, height * 0.22, width * 0.82, height * 0.68),
            (width * 0.62, height * 0.34, width * 0.90, height * 0.72),
        ]
        for circle in circles:
            draw.ellipse(circle, fill=0)
        draw.rounded_rectangle((width * 0.18, height * 0.48, width * 0.82, height * 0.82), radius=max(18, padding), fill=0)

    mask = np.array(image)
    if np.all(mask == 255):
        raise ValueError(f"Built-in shape '{shape_name}' did not produce a usable mask.")
    return mask


def load_mask_from_png(mask_path: str, width: int, height: int) -> np.ndarray:
    try:
        with Image.open(mask_path) as source_image:
            rgba = source_image.convert("RGBA").resize((width, height))
    except Exception as exc:
        raise ValueError(f"Failed to read PNG mask '{mask_path}': {exc}") from exc

    alpha = np.array(rgba.getchannel("A"), dtype=np.uint8)
    rgb = np.array(rgba.convert("RGB"), dtype=np.uint8)
    luminance = (rgb[:, :, 0].astype(np.float32) * 0.299) + (rgb[:, :, 1].astype(np.float32) * 0.587) + (rgb[:, :, 2].astype(np.float32) * 0.114)

    drawable = (alpha > 15) & (luminance < 245)
    if not np.any(drawable):
        raise ValueError("The selected PNG mask is fully transparent or too close to white to define a shape.")

    mask = np.where(drawable, 0, 255).astype(np.uint8)
    return mask


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


def get_default_visual_config() -> WordCloudConfig:
    return get_template_config(DEFAULT_PRESET_NAME)


def get_preset_store_path() -> str:
    home = Path.home()
    preset_dir = home / ".text_analyzer_pro"
    preset_dir.mkdir(parents=True, exist_ok=True)
    return str(preset_dir / "wordcloud_presets.json")


def serialize_preset_config(config: WordCloudConfig) -> Dict[str, object]:
    payload = asdict(config)
    payload["custom_stopwords"] = sorted(config.custom_stopwords)
    payload["include_terms"] = sorted(config.include_terms)
    payload["exclude_terms"] = sorted(config.exclude_terms)
    return payload


def deserialize_preset_config(payload: Dict[str, object]) -> WordCloudConfig:
    if not isinstance(payload, dict):
        raise ValueError("Preset data must be a JSON object.")
    return WordCloudConfig(**payload)


def load_presets() -> Dict[str, Dict[str, object]]:
    path = get_preset_store_path()
    if not os.path.exists(path):
        return {}

    with open(path, "r", encoding="utf-8") as handle:
        raw_data = json.load(handle)

    if not isinstance(raw_data, dict):
        raise ValueError("Preset file is malformed.")

    presets: Dict[str, Dict[str, object]] = {}
    for preset_name, payload in raw_data.items():
        if not isinstance(preset_name, str) or not preset_name.strip():
            raise ValueError("Preset names must be non-empty strings.")
        config = deserialize_preset_config(payload)
        presets[preset_name.strip()] = serialize_preset_config(config)
    return presets


def save_presets(presets: Dict[str, Dict[str, object]]) -> str:
    normalized: Dict[str, Dict[str, object]] = {}
    for preset_name, payload in presets.items():
        clean_name = str(preset_name).strip()
        if not clean_name:
            raise ValueError("Preset names must be non-empty strings.")
        config = deserialize_preset_config(payload)
        normalized[clean_name] = serialize_preset_config(config)

    path = get_preset_store_path()
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(normalized, handle, indent=2, sort_keys=True)
    return path


def save_preset(name: str, config: WordCloudConfig) -> str:
    clean_name = str(name).strip()
    if not clean_name:
        raise ValueError("Preset name cannot be empty.")
    presets = load_presets()
    presets[clean_name] = serialize_preset_config(config)
    return save_presets(presets)


def delete_preset(name: str) -> str:
    clean_name = str(name).strip()
    presets = load_presets()
    if clean_name not in presets:
        raise ValueError(f"Preset '{clean_name}' does not exist.")
    del presets[clean_name]
    return save_presets(presets)


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


def _apply_term_filters(stats_df: pd.DataFrame, config: WordCloudConfig) -> pd.DataFrame:
    if stats_df.empty:
        return stats_df

    filtered = stats_df.copy()
    normalized_terms = filtered["term"].astype(str).str.lower()

    if config.include_terms:
        filtered = filtered.loc[normalized_terms.isin(config.include_terms)].copy()
        normalized_terms = filtered["term"].astype(str).str.lower()

    if config.exclude_terms:
        filtered = filtered.loc[~normalized_terms.isin(config.exclude_terms)].copy()

    if filtered.empty:
        return pd.DataFrame(columns=["term", "count", "share"])

    if config.sort_mode == "Alphabetical":
        filtered = filtered.sort_values(["term", "count"], ascending=[True, False], kind="stable")
    else:
        filtered = filtered.sort_values(["count", "term"], ascending=[False, True], kind="stable")

    if config.render_top_n:
        filtered = filtered.head(config.render_top_n).copy()

    kept_total = filtered["count"].sum()
    filtered["share"] = filtered["count"] / kept_total if kept_total else 0
    return filtered.reset_index(drop=True)


def _normalize_stopword(word: str) -> str:
    return str(word).strip().lower()


def _normalize_color_value(value: str) -> str:
    text = str(value).strip()
    if not text:
        raise ValueError("Color values cannot be empty.")
    try:
        ImageColor.getrgb(text)
    except ValueError as exc:
        raise ValueError(f"Invalid color value: {value}") from exc
    return text


def _normalize_hex_color(value: str) -> str:
    text = str(value).strip()
    if not HEX_COLOR_RE.match(text):
        raise ValueError(f"Invalid hex color: {value}")
    if len(text) == 4:
        text = "#" + "".join(char * 2 for char in text[1:])
    return text.lower()


# ============================================================================
# Distributable-compatible wordcloud API
# ============================================================================

# Gradient color schemes (matplotlib colormaps)
WORDCLOUD_GRADIENT_SCHEMES = {
    "Corporate Blue": "Blues",
    "Navy Gradient": "PuBu",
    "Slate": "Greys",
    "Ocean": "GnBu",
    "Forest": "Greens",
    "Earth Tones": "YlOrBr",
    "Lavender": "Purples",
    "Sunset": "YlOrRd",
    "Fire": "OrRd",
    "Berry": "RdPu",
    "Viridis": "viridis",
    "Plasma": "plasma",
    "Magma": "magma",
    "Inferno": "inferno",
    "Cool": "cool",
    "Warm": "autumn",
    "Grayscale": "gray",
}

WORDCLOUD_MULTI_COLOR_PALETTES = {
    "Rainbow Mix": ["#e74c3c", "#f39c12", "#f1c40f", "#2ecc71", "#3498db", "#9b59b6"],
    "Candy": ["#ff6b6b", "#feca57", "#48dbfb", "#ff9ff3", "#54a0ff", "#5f27cd"],
    "Neon": ["#39ff14", "#ff073a", "#00f7ff", "#ff00ff", "#ffff00", "#ff6600"],
    "Pastel Mix": ["#a8e6cf", "#dcedc1", "#ffd3b6", "#ffaaa5", "#d5aaff", "#a0c4ff"],
    "Primary Colors": ["#e74c3c", "#3498db", "#f1c40f", "#2ecc71"],
    "Ocean Breeze": ["#0077b6", "#00b4d8", "#90e0ef", "#48cae4", "#023e8a"],
    "Autumn Leaves": ["#d4a373", "#e07a5f", "#f2cc8f", "#81b29a", "#3d405b"],
    "Berry Blast": ["#7209b7", "#b5179e", "#f72585", "#3a0ca3", "#560bad"],
    "Mint & Coral": ["#00b894", "#00cec9", "#fab1a0", "#e17055", "#81ecec"],
    "Sunset Beach": ["#ff6b35", "#f7c59f", "#004e89", "#1a659e", "#ff9f1c"],
    "Forest Floor": ["#2d6a4f", "#40916c", "#52b788", "#74c69d", "#95d5b2"],
    "Vintage": ["#6b705c", "#a5a58d", "#b7b7a4", "#ffe8d6", "#ddbea9"],
    "Tech": ["#00d4ff", "#7928ca", "#ff0080", "#00ff88", "#ff6b6b"],
    "Earthy": ["#bc6c25", "#dda15e", "#606c38", "#283618", "#fefae0"],
}

WORDCLOUD_COLOR_SCHEMES = {
    **WORDCLOUD_GRADIENT_SCHEMES,
    **WORDCLOUD_MULTI_COLOR_PALETTES,
}

WORDCLOUD_BACKGROUNDS = {
    "White": "white",
    "Off-White": "#f8f8f8",
    "Light Gray": "#e0e0e0",
    "Dark Gray": "#2d2d2d",
    "Black": "black",
    "Navy": "#1a1a2e",
    "Cream": "#fffef0",
    "Light Blue": "#e3f2fd",
    "Transparent": None,
}

WORDCLOUD_SHAPES = {
    "Rectangle": "rectangle",
    "Circle": "circle",
    "Oval": "oval",
    "Rounded Rectangle": "rounded_rect",
    "Diamond": "diamond",
    "Heart": "heart",
    "Star": "star",
    "Cloud": "cloud",
    "Hexagon": "hexagon",
    "Triangle": "triangle",
}


def create_shape_mask(shape: str, width: int = 1920, height: int = 1080) -> Optional[np.ndarray]:
    """Create a mask array for the given shape (distributable-compatible).

    Args:
        shape: Shape value from WORDCLOUD_SHAPES (e.g. "circle", "heart")
        width: Image width
        height: Image height

    Returns:
        Numpy array mask (0 = word area, 255 = excluded) or None for rectangle.
    """
    if shape == "rectangle" or shape is None:
        return None

    mask = np.ones((height, width), dtype=np.uint8) * 255
    cx, cy = width // 2, height // 2
    rx, ry = width // 2 - 50, height // 2 - 50

    if shape == "circle":
        r = min(rx, ry)
        y, x = np.ogrid[:height, :width]
        dist = np.sqrt((x - cx) ** 2 + (y - cy) ** 2)
        mask[dist <= r] = 0
    elif shape == "oval":
        y, x = np.ogrid[:height, :width]
        dist = ((x - cx) / rx) ** 2 + ((y - cy) / ry) ** 2
        mask[dist <= 1] = 0
    elif shape == "rounded_rect":
        corner_r = min(rx, ry) // 4
        mask[cy - ry + corner_r : cy + ry - corner_r, cx - rx : cx + rx] = 0
        mask[cy - ry : cy + ry, cx - rx + corner_r : cx + rx - corner_r] = 0
        for dx, dy in [(-1, -1), (-1, 1), (1, -1), (1, 1)]:
            corner_cx = cx + dx * (rx - corner_r)
            corner_cy = cy + dy * (ry - corner_r)
            y, x = np.ogrid[:height, :width]
            dist = np.sqrt((x - corner_cx) ** 2 + (y - corner_cy) ** 2)
            mask[dist <= corner_r] = 0
    elif shape == "diamond":
        y, x = np.ogrid[:height, :width]
        dist = np.abs(x - cx) / rx + np.abs(y - cy) / ry
        mask[dist <= 1] = 0
    elif shape == "heart":
        y, x = np.ogrid[:height, :width]
        xn = (x - cx) / (rx * 0.8)
        yn = (cy - y) / (ry * 0.9)
        heart = (xn**2 + yn**2 - 1) ** 3 - xn**2 * yn**3
        mask[heart <= 0] = 0
    elif shape == "star":
        points_outer, points_inner = [], []
        for i in range(5):
            angle_outer = math.radians(90 + i * 72)
            angle_inner = math.radians(90 + i * 72 + 36)
            points_outer.append((cx + int(rx * 0.95 * math.cos(angle_outer)),
                                 cy - int(ry * 0.95 * math.sin(angle_outer))))
            points_inner.append((cx + int(rx * 0.4 * math.cos(angle_inner)),
                                 cy - int(ry * 0.4 * math.sin(angle_inner))))
        star_points = []
        for i in range(5):
            star_points.append(points_outer[i])
            star_points.append(points_inner[i])
        _fill_shape_polygon(mask, star_points, 0)
    elif shape == "cloud":
        circles = [
            (cx - rx * 0.5, cy, min(rx, ry) * 0.5),
            (cx + rx * 0.5, cy, min(rx, ry) * 0.5),
            (cx, cy - ry * 0.3, min(rx, ry) * 0.55),
            (cx - rx * 0.25, cy - ry * 0.15, min(rx, ry) * 0.45),
            (cx + rx * 0.25, cy - ry * 0.15, min(rx, ry) * 0.45),
            (cx, cy + ry * 0.2, min(rx, ry) * 0.4),
        ]
        y, x = np.ogrid[:height, :width]
        for ccx, ccy, cr in circles:
            dist = np.sqrt((x - ccx) ** 2 + (y - ccy) ** 2)
            mask[dist <= cr] = 0
    elif shape == "hexagon":
        points = []
        for i in range(6):
            angle = math.radians(60 * i)
            points.append((cx + int(rx * math.cos(angle)),
                           cy + int(ry * math.sin(angle))))
        _fill_shape_polygon(mask, points, 0)
    elif shape == "triangle":
        points = [
            (cx, cy - int(ry * 0.95)),
            (cx - int(rx * 0.95), cy + int(ry * 0.7)),
            (cx + int(rx * 0.95), cy + int(ry * 0.7)),
        ]
        _fill_shape_polygon(mask, points, 0)

    return mask


def _fill_shape_polygon(mask: np.ndarray, points, value: int):
    """Fill a polygon in the mask using PIL scanline."""
    height, width = mask.shape
    img = Image.new("L", (width, height), 255)
    draw = ImageDraw.Draw(img)
    draw.polygon(points, fill=value)
    mask[:] = np.array(img)


def _bg_luminance(color) -> float:
    """Perceived luminance (0..1) for a CSS color string. Treats transparent
    / None as light (most decks use light slides)."""
    if color is None or color == "transparent":
        return 1.0
    try:
        rgb = ImageColor.getrgb(color)[:3]
    except (ValueError, TypeError):
        return 1.0
    r, g, b = (c / 255.0 for c in rgb)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _create_contrast_aware_gradient_color_func(colormap_name: str, background_color):
    """Build a color_func that samples a matplotlib colormap, but only within
    a sub-range that contrasts against the chosen background.

    WordCloud's default `colormap=...` uniformly samples [0..1], so on a white
    background the low end of e.g. ``Blues`` produces near-invisible pale text.
    This helper clips the sampling range so even the rarest words read clearly.
    """
    import matplotlib.cm as mpl_cm

    cmap = mpl_cm.get_cmap(colormap_name)
    lum = _bg_luminance(background_color)
    if lum > 0.65:
        # Light background: keep darker end of the gradient.
        low, high = 0.45, 1.0
    elif lum < 0.35:
        # Dark background: keep brighter end.
        low, high = 0.0, 0.55
    else:
        # Medium: compress both ends slightly.
        low, high = 0.25, 0.9

    def color_func(word, font_size, position, orientation, random_state=None, **kwargs):
        if random_state is None:
            random_state = np.random.RandomState()
        t = random_state.uniform(low, high)
        r, g, b = (int(c * 255) for c in cmap(t)[:3])
        return f"rgb({r}, {g}, {b})"

    return color_func


def _create_wc_color_func(palette):
    """Create a color function that randomly selects from a palette list."""
    def color_func(word, font_size, position, orientation, random_state=None, **kwargs):
        return random.choice(palette)
    return color_func


def generate_wordcloud(
    texts,
    colormap="Blues",
    background_color="white",
    max_words=200,
    width=1920,
    height=1080,
    mask=None,
    font_path=None,
    min_font_size=10,
    max_font_size=None,
    stopwords=None,
    relative_scaling=0.0,
):
    """Generate a word cloud (distributable-compatible API).

    Returns:
        Tuple of (WordCloud object, word_frequencies dict).
        Returns (None, {}) if wordcloud is not available.
    """
    if WordCloud is None:
        return None, {}

    if stopwords is None:
        stopwords = set(ENGLISH_STOP_WORDS)

    stopwords_lower = {w.lower() for w in stopwords}
    word_counts = {}
    for text in texts:
        for word in re.findall(r"\b[a-zA-Z0-9]+\b", str(text).lower()):
            if word not in stopwords_lower and len(word) > 1:
                word_counts[word] = word_counts.get(word, 0) + 1

    if not word_counts:
        return None, {}

    is_multi_color = isinstance(colormap, list)
    color_func = None
    cmap = None

    if is_multi_color:
        color_func = _create_wc_color_func(colormap)
    elif colormap:
        # Use a contrast-aware sampler instead of WordCloud's full-range
        # uniform sampling so low-frequency words don't disappear into the
        # background.
        color_func = _create_contrast_aware_gradient_color_func(colormap, background_color)

    wc = WordCloud(
        width=width,
        height=height,
        max_words=max_words,
        colormap=cmap,
        background_color=background_color,
        mask=mask,
        font_path=font_path,
        min_font_size=min_font_size,
        max_font_size=max_font_size,
        stopwords=stopwords,
        mode="RGBA" if background_color is None else "RGB",
        prefer_horizontal=0.9,
        relative_scaling=relative_scaling,
    )

    wc.generate_from_frequencies(word_counts)

    if color_func:
        wc.recolor(color_func=color_func)

    word_frequencies = wc.words_
    return wc, word_frequencies


def wordcloud_to_image(wc):
    """Convert WordCloud to PIL Image."""
    if wc is None:
        return None
    try:
        return wc.to_image()
    except Exception:
        return None


def save_wordcloud(wc, path, dpi=300):
    """Save word cloud to file (PNG, JPG, or SVG)."""
    if wc is None:
        return False
    try:
        if path.lower().endswith((".png", ".jpg", ".jpeg")):
            img = wc.to_image()
            img.save(path, dpi=(dpi, dpi))
        elif path.lower().endswith(".svg"):
            import matplotlib.pyplot as plt

            fig, ax = plt.subplots(figsize=(19.2, 10.8))
            ax.imshow(wc, interpolation="bilinear")
            ax.axis("off")
            fig.savefig(path, format="svg", bbox_inches="tight", pad_inches=0)
            plt.close(fig)
        else:
            img = wc.to_image()
            img.save(path, dpi=(dpi, dpi))
        return True
    except Exception:
        return False


def load_custom_mask(image_path):
    """Load an image file and convert it to a word cloud mask array."""
    try:
        img = Image.open(image_path).convert("RGBA")
        img.thumbnail((1920, 1080), Image.Resampling.LANCZOS)
        canvas = Image.new("RGBA", (1920, 1080), (255, 255, 255, 255))
        offset = ((1920 - img.width) // 2, (1080 - img.height) // 2)
        canvas.paste(img, offset, img)
        arr = np.array(canvas)
        alpha = arr[:, :, 3]
        mask = np.where(alpha > 128, 0, 255).astype(np.uint8)
        if mask.min() == mask.max():
            gray = np.array(canvas.convert("L"))
            mask = np.where(gray < 200, 0, 255).astype(np.uint8)
        if mask.min() == mask.max():
            return None
        return mask
    except Exception:
        return None


def discover_system_fonts():
    """Discover available TrueType fonts from the system Fonts directory."""
    fonts = {"(Default)": None}
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    if not os.path.isdir(fonts_dir):
        return fonts
    try:
        for f in sorted(os.listdir(fonts_dir)):
            if f.lower().endswith((".ttf", ".otf")):
                name = os.path.splitext(f)[0]
                for suffix in (
                    "-Regular", "-Bold", "-Italic", "-BoldItalic",
                    "Regular", "Bold", "Italic",
                ):
                    if name.endswith(suffix):
                        name = name[: -len(suffix)].rstrip("-_ ")
                        break
                if name:
                    full_path = os.path.join(fonts_dir, f)
                    fonts[name] = full_path
    except OSError:
        pass
    return fonts
