"""Design tokens for Text Analyzer Pro themes.

Two palettes (light, dark) with shared spacing / radius / type tokens.
Keep colors compact — referenced by name from qss.build_qss().

This module is the single source of truth for theme colors, spacing,
radius, typography, and platform-aware fonts.
"""

from __future__ import annotations

import sys

LIGHT = {
    "name": "light",
    # surfaces
    "bg": "#f5f6f8",
    "bg_elev": "#ffffff",
    "surface": "#ffffff",
    "surface_alt": "#f0f2f5",
    "sidebar_bg": "#1e2230",
    "sidebar_fg": "#cfd3dc",
    "sidebar_fg_active": "#ffffff",
    "sidebar_active_bg": "#2a3145",
    "sidebar_accent": "#4f8cff",
    # borders / dividers
    "border": "#dfe3eb",
    "border_strong": "#c5cad3",
    "divider": "#eaecef",
    # text
    "text": "#1d2230",
    "text_muted": "#5b6271",
    "text_subtle": "#8a90a0",
    "text_inverse": "#ffffff",
    # accents
    "accent": "#3d6dff",
    "accent_hover": "#2f5be0",
    "accent_pressed": "#264bbf",
    "accent_soft": "#e7eeff",
    "accent_subtle": "#f1f5ff",
    # state
    "success": "#1f9d55",
    "success_soft": "#e2f5ea",
    "warning": "#d29922",
    "warning_soft": "#fbf2dc",
    "danger": "#d73a49",
    "danger_hover": "#b22a36",
    "danger_soft": "#fbe4e6",
    "info": "#1f8fff",
    "info_soft": "#e2f0ff",
    # input
    "input_bg": "#ffffff",
    "input_border": "#cfd3dc",
    "input_focus": "#3d6dff",
    "placeholder": "#9aa0ac",
    # selection
    "selection_bg": "#dbe5ff",
    "selection_fg": "#1d2230",
    # tooltip
    "tooltip_bg": "#1d2230",
    "tooltip_fg": "#ffffff",
    # scrollbar
    "scroll_bg": "#eef0f3",
    "scroll_thumb": "#c2c7d0",
    "scroll_thumb_hover": "#9aa0ac",
    # focus / shadow
    "focus_ring": "#3d6dff",
    "shadow_color": "rgba(15, 23, 42, 0.10)",
    "shadow_color_strong": "rgba(15, 23, 42, 0.18)",
    # KBD chip
    "kbd_bg": "#f0f2f5",
    "kbd_fg": "#5b6271",
    "kbd_border": "#dfe3eb",
}

DARK = {
    "name": "dark",
    "bg": "#15171c",
    "bg_elev": "#1c1f26",
    "surface": "#22262e",
    "surface_alt": "#2a2f38",
    "sidebar_bg": "#101218",
    "sidebar_fg": "#a8aebc",
    "sidebar_fg_active": "#ffffff",
    "sidebar_active_bg": "#1f2638",
    "sidebar_accent": "#5b8dff",
    "border": "#2f343d",
    "border_strong": "#3b4150",
    "divider": "#262a32",
    "text": "#e6e8ee",
    "text_muted": "#b3b8c4",
    "text_subtle": "#7e8494",
    "text_inverse": "#15171c",
    "accent": "#5b8dff",
    "accent_hover": "#7aa3ff",
    "accent_pressed": "#3d6dff",
    "accent_soft": "#1f2a44",
    "accent_subtle": "#1a2236",
    "success": "#3bbd72",
    "success_soft": "#16321f",
    "warning": "#e8b347",
    "warning_soft": "#3a2c10",
    "danger": "#e5534b",
    "danger_hover": "#cc433c",
    "danger_soft": "#3a1c1d",
    "info": "#5fb1ff",
    "info_soft": "#15263a",
    "input_bg": "#1a1d24",
    "input_border": "#343a47",
    "input_focus": "#5b8dff",
    "placeholder": "#7e8494",
    "selection_bg": "#2a3a64",
    "selection_fg": "#ffffff",
    "tooltip_bg": "#0f1115",
    "tooltip_fg": "#e6e8ee",
    "scroll_bg": "#1a1d24",
    "scroll_thumb": "#3b4150",
    "scroll_thumb_hover": "#525a6c",
    "focus_ring": "#7aa3ff",
    "shadow_color": "rgba(0, 0, 0, 0.45)",
    "shadow_color_strong": "rgba(0, 0, 0, 0.65)",
    "kbd_bg": "#22262e",
    "kbd_fg": "#b3b8c4",
    "kbd_border": "#343a47",
}


def _platform_font_family() -> str:
    """Return a CSS-style font stack tuned for the host OS.

    The first families are platform-preferred; the remainder are fallbacks
    so the app stays readable on any platform.
    """
    if sys.platform == "darwin":
        first = "'SF Pro Text', '-apple-system', 'Helvetica Neue'"
    elif sys.platform.startswith("win"):
        first = "'Segoe UI Variable', 'Segoe UI', 'Inter'"
    else:
        first = "'Inter', 'Ubuntu', 'Cantarell', 'Noto Sans'"
    fallback = "'system-ui', 'Helvetica Neue', Arial, sans-serif"
    return f"{first}, {fallback}"


def _platform_font_size_px() -> int:
    """Return base font size in pixels tuned per platform."""
    # Keep all platforms at 13px for now — predictable rendering across OSes;
    # platform-specific tweaks can land later if needed.
    return 13


_BASE = _platform_font_size_px()

# Shared (non-color) tokens
SHARED = {
    # radius
    "radius_sm": "4px",
    "radius_md": "8px",
    "radius_lg": "12px",
    "radius_xl": "16px",
    # spacing (4 / 8 / 12 / 16 / 24)
    "space_xs": "4px",
    "space_sm": "8px",
    "space_md": "12px",
    "space_lg": "16px",
    "space_xl": "24px",
    # typography
    "font_family": _platform_font_family(),
    "font_mono": "'Cascadia Mono', 'SF Mono', 'JetBrains Mono', 'Consolas', 'Menlo', monospace",
    "font_size_xs": f"{max(_BASE - 3, 9)}px",
    "font_size_sm": f"{max(_BASE - 2, 10)}px",
    "font_size_md": f"{_BASE}px",
    "font_size_lg": f"{_BASE + 2}px",
    "font_size_xl": f"{_BASE + 5}px",
    "font_size_xxl": f"{_BASE + 9}px",
    # base font as plain int (used by ThemeManager to set QFont)
    "_base_font_pt": _BASE,
    # elevation z-levels (numeric blur radii used by effects.apply_card_shadow)
    "_elevation_1_blur": 8,
    "_elevation_2_blur": 16,
    "_elevation_3_blur": 28,
}


def get_tokens(mode: str) -> dict:
    """Return merged token dict for the requested mode ('light' or 'dark')."""
    base = DARK if str(mode).lower() == "dark" else LIGHT
    merged = {}
    merged.update(SHARED)
    merged.update(base)
    return merged
