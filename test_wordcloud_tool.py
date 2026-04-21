import os
import tempfile
import unittest
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


class FakeApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.app_title = "Test App"
        self.df = pd.DataFrame({"text": ["alpha beta", "beta gamma"]})
        self.current_file_path = os.path.join(os.getcwd(), "demo.xlsx")
        self.wordcloud_builder = None

    def current_column_name(self):
        return "text"

    def current_sheet_name(self):
        return "Sheet1"

    def log_msg(self, _message):
        return None


class WordCloudBuilderGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qapp = get_qapp()

    def setUp(self):
        self.app = FakeApp()
        self.builder = WordCloudBuilderWindow(self.app)

    def tearDown(self):
        self.builder.close()
        self.app.close()

    def test_visual_control_state_switches(self):
        self.builder.color_mode_combo.setCurrentText("Custom")
        self.builder.mask_mode_combo.setCurrentText("Custom PNG")
        self.builder._update_visual_control_states()
        self.assertTrue(self.builder.custom_colors_edit.isEnabled())
        self.assertTrue(self.builder.mask_path_edit.isEnabled())
        self.assertFalse(self.builder.palette_combo.isEnabled())

    def test_reset_to_default_restores_visual_defaults(self):
        self.builder.color_mode_combo.setCurrentText("Custom")
        self.builder.custom_colors_edit.setText("#123456")
        self.builder.mask_mode_combo.setCurrentText("Builtin Shape")
        self.builder.shape_combo.setCurrentText("Heart")
        self.builder.reset_to_default_preset()
        self.assertEqual(self.builder.color_mode_combo.currentText(), "Colormap")
        self.assertEqual(self.builder.mask_mode_combo.currentText(), "None")
        self.assertEqual(self.builder.font_label.text(), "Default font")

    def test_live_preview_is_scheduled_on_setting_change(self):
        with mock.patch.object(self.builder, "schedule_live_preview") as schedule_preview:
            self.builder._on_live_setting_changed()
        schedule_preview.assert_called()

    def test_template_selection_updates_form(self):
        self.builder.template_apply_combo.setCurrentText("High Contrast")
        self.builder.apply_selected_template()
        self.assertEqual(self.builder.active_template_label.text(), "High Contrast")
        self.assertEqual(self.builder.color_mode_combo.currentText(), "Palette")
        self.assertEqual(self.builder.mask_mode_combo.currentText(), "Builtin Shape")

    def test_font_selection_updates_font_path_and_label(self):
        if len(self.builder.FONT_OPTIONS) < 2:
            self.skipTest("No system fonts discovered in this environment")
        label, path = self.builder.FONT_OPTIONS[1]
        self.builder.font_choice_combo.setCurrentText(label)
        self.builder.apply_selected_font()
        self.assertEqual(self.builder.font_label.text(), label)
        self.assertEqual(getattr(self.builder, "_custom_font_path", ""), path)

    def test_stale_render_result_is_ignored(self):
        stats_df = pd.DataFrame([{"term": "alpha", "count": 2, "share": 1.0}])
        image = Image.new("RGB", (100, 80), "white")
        self.builder._latest_request_id = 2
        self.builder.is_rendering = True
        self.builder._finish_render(1, "text", stats_df, {"total_rows": 2, "usable_rows": 1, "unique_terms": 1, "kept_term_occurrences": 2}, image)
        self.assertIsNone(self.builder.current_image)


if __name__ == "__main__":
    unittest.main()
