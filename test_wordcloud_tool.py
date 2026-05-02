import os
import tempfile
import unittest
from collections import Counter
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pandas as pd
from PIL import Image
from PySide6 import QtWidgets

import wordcloud_tool
from gui import WordCloudBuilderWindow
from wordcloud_tool import (
    WordCloudConfig,
    build_builtin_mask,
    build_color_func,
    build_term_stats,
    delete_preset,
    export_term_stats,
    get_font_choices,
    get_template_config,
    get_template_names,
    load_mask_from_png,
    load_presets,
    render_wordcloud,
    save_preset,
    summarize_texts,
)


def get_qapp():
    app = QtWidgets.QApplication.instance()
    if app is None:
        app = QtWidgets.QApplication([])
    return app


class WordCloudToolTests(unittest.TestCase):
    def test_empty_column_returns_empty_stats_and_zero_summary(self):
        config = WordCloudConfig()
        stats_df = build_term_stats(["", None, "   "], config)
        summary = summarize_texts(["", None, "   "], config)
        self.assertEqual(list(stats_df.columns), ["term", "count", "share"])
        self.assertTrue(stats_df.empty)
        self.assertEqual(summary["total_rows"], 3)
        self.assertEqual(summary["usable_rows"], 0)
        self.assertEqual(summary["unique_terms"], 0)
        self.assertEqual(summary["kept_term_occurrences"], 0)

    def test_builtin_and_custom_stopwords_are_removed(self):
        config = WordCloudConfig(custom_stopwords={"noise", "beta"})
        stats_df = build_term_stats(["The alpha beta and beta", "ALPHA noise"], config)
        self.assertEqual(stats_df["term"].tolist(), ["alpha"])
        self.assertEqual(int(stats_df.iloc[0]["count"]), 2)

    def test_numeric_only_tokens_can_be_excluded(self):
        excluded_stats = build_term_stats(["alpha 2024 123 alpha", "2024 beta"], WordCloudConfig(exclude_numeric=True))
        included_stats = build_term_stats(["alpha 2024 123 alpha", "2024 beta"], WordCloudConfig(exclude_numeric=False))
        self.assertNotIn("2024", excluded_stats["term"].tolist())
        self.assertIn("2024", included_stats["term"].tolist())

    def test_phrase_mode_includes_expected_ngrams(self):
        texts = ["alpha beta gamma"]
        unigram_terms = build_term_stats(texts, WordCloudConfig(phrase_mode="Unigrams"))["term"].tolist()
        bigram_terms = build_term_stats(texts, WordCloudConfig(phrase_mode="Up to Bigrams"))["term"].tolist()
        trigram_terms = build_term_stats(texts, WordCloudConfig(phrase_mode="Up to Trigrams"))["term"].tolist()
        self.assertEqual(unigram_terms, ["alpha", "beta", "gamma"])
        self.assertIn("alpha beta", bigram_terms)
        self.assertIn("alpha beta gamma", trigram_terms)

    def test_min_frequency_filters_and_preserves_stable_sorting(self):
        stats_df = build_term_stats(
            ["banana apple carrot", "banana apple", "banana pear", "apple pear"],
            WordCloudConfig(min_frequency=2),
        )
        self.assertEqual(stats_df["term"].tolist(), ["apple", "banana", "pear"])
        self.assertEqual(stats_df["count"].tolist(), [3, 3, 2])

    def test_wordcloud_config_validates_new_ranges(self):
        with self.assertRaises(ValueError):
            WordCloudConfig(prefer_horizontal=1.5)
        with self.assertRaises(ValueError):
            WordCloudConfig(relative_scaling=-0.1)
        with self.assertRaises(ValueError):
            WordCloudConfig(scale=0)
        with self.assertRaises(ValueError):
            WordCloudConfig(color_mode="Custom", custom_colors=[])

    def test_custom_palette_color_func_is_created(self):
        config = WordCloudConfig(color_mode="Custom", custom_colors=["#123456", "#abcdef"])
        color_func = build_color_func(config)
        self.assertIsNotNone(color_func)
        self.assertIn(color_func("alpha", 24, (0, 0), None), {"#123456", "#abcdef"})

    def test_builtin_mask_generation_returns_non_empty_mask(self):
        mask = build_builtin_mask("Heart", 320, 200)
        self.assertEqual(mask.shape, (200, 320))
        self.assertTrue((mask == 0).any())

    def test_new_builtin_shapes_generate_masks(self):
        for shape_name in ("Diamond", "Hexagon", "Triangle", "Shield", "Cloud"):
            mask = build_builtin_mask(shape_name, 320, 200)
            self.assertEqual(mask.shape, (200, 320))
            self.assertTrue((mask == 0).any(), shape_name)

    def test_font_choices_include_default(self):
        choices = get_font_choices()
        self.assertGreaterEqual(len(choices), 1)
        self.assertEqual(choices[0][0], "Default font")

    def test_load_mask_from_png_rejects_blank_masks(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            mask_path = os.path.join(tmpdir, "blank.png")
            Image.new("RGBA", (40, 30), (255, 255, 255, 0)).save(mask_path)
            with self.assertRaises(ValueError):
                load_mask_from_png(mask_path, 80, 60)

    def test_load_mask_from_png_accepts_drawable_area(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            mask_path = os.path.join(tmpdir, "shape.png")
            image = Image.new("RGBA", (40, 30), (255, 255, 255, 0))
            for x in range(10, 30):
                for y in range(8, 22):
                    image.putpixel((x, y), (0, 0, 0, 255))
            image.save(mask_path)
            mask = load_mask_from_png(mask_path, 80, 60)
        self.assertEqual(mask.shape, (60, 80))
        self.assertTrue((mask == 0).any())

    def test_preset_round_trip_and_delete(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            preset_path = os.path.join(tmpdir, "presets.json")
            config = WordCloudConfig(color_mode="Custom", custom_colors=["#123456"], mask_mode="Builtin Shape", shape_name="Star")
            with mock.patch("wordcloud_tool.get_preset_store_path", return_value=preset_path):
                save_preset("Demo", config)
                loaded = load_presets()
                self.assertIn("Demo", loaded)
                self.assertEqual(loaded["Demo"]["shape_name"], "Star")
                delete_preset("Demo")
                self.assertEqual(load_presets(), {})

    def test_malformed_preset_file_is_rejected(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            preset_path = os.path.join(tmpdir, "presets.json")
            with open(preset_path, "w", encoding="utf-8") as handle:
                handle.write('{"Broken": {"color_mode": "Custom", "custom_colors": []}}')
            with mock.patch("wordcloud_tool.get_preset_store_path", return_value=preset_path):
                with self.assertRaises(ValueError):
                    load_presets()

    def test_include_exclude_and_top_n_filters_are_applied(self):
        stats_df = build_term_stats(
            ["alpha beta gamma", "alpha beta", "alpha delta"],
            WordCloudConfig(include_terms={"alpha", "beta"}, exclude_terms={"beta"}, render_top_n=1),
        )
        self.assertEqual(stats_df["term"].tolist(), ["alpha"])
        self.assertEqual(stats_df["count"].tolist(), [3])

    def test_template_catalog_returns_expected_template(self):
        self.assertIn("Executive Clean", get_template_names())
        config = get_template_config("Executive Clean")
        self.assertEqual(config.template_name, "Executive Clean")
        self.assertEqual(config.color_mode, "Palette")

    def test_deserialize_backward_compatible_without_new_fields(self):
        config = WordCloudConfig(**{"max_words": 150, "palette_name": "Viridis"})
        self.assertEqual(config.render_top_n, 0)
        self.assertEqual(config.include_terms, set())
        self.assertEqual(config.exclude_terms, set())

    def test_export_term_stats_writes_excel_output(self):
        stats_df = build_term_stats(["alpha beta alpha"], WordCloudConfig())
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = f"{tmpdir}/terms.xlsx"
            saved_path = export_term_stats(stats_df, out_path)
            loaded = pd.read_excel(saved_path, engine="openpyxl")
        self.assertEqual(saved_path, out_path)
        self.assertEqual(loaded["term"].tolist(), ["alpha", "beta"])

    def test_render_wordcloud_requires_dependency_or_returns_image(self):
        stats_df = build_term_stats(["alpha beta alpha"], WordCloudConfig())
        config = WordCloudConfig(width=320, height=200)
        if wordcloud_tool.WordCloud is None:
            with self.assertRaises(ImportError):
                render_wordcloud(stats_df, config)
            return
        image = render_wordcloud(stats_df, config)
        self.assertEqual(image.size, (320, 200))

    def test_render_wordcloud_supports_masks_and_custom_colors(self):
        stats_df = build_term_stats(["alpha beta alpha gamma"], WordCloudConfig())
        config = WordCloudConfig(width=280, height=180, color_mode="Custom", custom_colors=["#264653", "#2a9d8f", "#e9c46a"], mask_mode="Builtin Shape", shape_name="Circle")
        if wordcloud_tool.WordCloud is None:
            with self.assertRaises(ImportError):
                render_wordcloud(stats_df, config)
            return
        image = render_wordcloud(stats_df, config)
        self.assertEqual(image.size, (280, 180))


class WordCloudDialogGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qapp = get_qapp()

    def setUp(self):
        self.parent = QtWidgets.QWidget()
        self.texts = ["alpha beta alpha", "beta gamma delta"]
        # Avoid mock.patch.object — PySide6 signal system can segfault with MagicMock slots.
        orig = WordCloudBuilderWindow._generate_preview
        WordCloudBuilderWindow._generate_preview = lambda self: None
        try:
            self.dialog = WordCloudBuilderWindow(self.parent, self.texts, "text")
        finally:
            WordCloudBuilderWindow._generate_preview = orig

    def tearDown(self):
        self.dialog.close()
        self.parent.close()

    def test_shape_combo_has_custom_image_option(self):
        items = [self.dialog.shape_combo.itemText(i)
                 for i in range(self.dialog.shape_combo.count())]
        self.assertIn("Custom Image\u2026", items)
        self.assertIn("Heart", items)
        self.assertIn("Circle", items)

    def test_rel_scale_slider_updates_label(self):
        self.dialog.rel_scale_slider.setValue(0)
        self.assertIn("rank only", self.dialog.rel_scale_label.text())
        self.dialog.rel_scale_slider.setValue(100)
        self.assertIn("fully proportional", self.dialog.rel_scale_label.text())

    def test_stopwords_display_updates(self):
        self.dialog.custom_stopwords = {"hello", "world"}
        self.dialog._update_stopwords_display()
        self.assertIn("2", self.dialog.stopwords_count_label.text())

    def test_word_filter_clears_tree(self):
        self.dialog.actual_word_counts = Counter({"alpha": 5, "beta": 3})
        self.dialog.total_word_count = 8
        self.dialog._update_word_counts()
        self.assertGreater(self.dialog.word_tree.topLevelItemCount(), 0)
        self.dialog.word_filter_edit.setText("zzz")
        self.dialog._on_word_filter_change()
        self.assertEqual(self.dialog.word_tree.topLevelItemCount(), 0)

    def test_color_combo_contains_distributable_schemes(self):
        items = [self.dialog.color_combo.itemText(i)
                 for i in range(self.dialog.color_combo.count())]
        self.assertIn("Corporate Blue", items)
        self.assertIn("Rainbow Mix", items)

    def test_background_combo_contains_all_options(self):
        items = [self.dialog.bg_combo.itemText(i)
                 for i in range(self.dialog.bg_combo.count())]
        self.assertIn("White", items)
        self.assertIn("Black", items)
        self.assertIn("Transparent", items)


if __name__ == "__main__":
    unittest.main()
