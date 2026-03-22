import tempfile
import unittest

import pandas as pd

import wordcloud_tool
from wordcloud_tool import (
    WordCloudConfig,
    build_term_stats,
    export_term_stats,
    render_wordcloud,
    summarize_texts,
)


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

        stats_df = build_term_stats(
            ["The alpha beta and beta", "ALPHA noise"],
            config,
        )

        self.assertEqual(stats_df["term"].tolist(), ["alpha"])
        self.assertEqual(int(stats_df.iloc[0]["count"]), 2)

    def test_numeric_only_tokens_can_be_excluded(self):
        excluded_stats = build_term_stats(
            ["alpha 2024 123 alpha", "2024 beta"],
            WordCloudConfig(exclude_numeric=True),
        )
        included_stats = build_term_stats(
            ["alpha 2024 123 alpha", "2024 beta"],
            WordCloudConfig(exclude_numeric=False),
        )

        self.assertNotIn("2024", excluded_stats["term"].tolist())
        self.assertNotIn("123", excluded_stats["term"].tolist())
        self.assertIn("2024", included_stats["term"].tolist())
        self.assertIn("123", included_stats["term"].tolist())

    def test_phrase_mode_includes_expected_ngrams(self):
        texts = ["alpha beta gamma"]

        unigram_terms = build_term_stats(texts, WordCloudConfig(phrase_mode="Unigrams"))["term"].tolist()
        bigram_terms = build_term_stats(texts, WordCloudConfig(phrase_mode="Up to Bigrams"))["term"].tolist()
        trigram_terms = build_term_stats(texts, WordCloudConfig(phrase_mode="Up to Trigrams"))["term"].tolist()

        self.assertEqual(unigram_terms, ["alpha", "beta", "gamma"])
        self.assertIn("alpha beta", bigram_terms)
        self.assertIn("beta gamma", bigram_terms)
        self.assertIn("alpha beta gamma", trigram_terms)

    def test_min_frequency_filters_and_preserves_stable_sorting(self):
        stats_df = build_term_stats(
            ["banana apple carrot", "banana apple", "banana pear", "apple pear"],
            WordCloudConfig(min_frequency=2),
        )

        self.assertEqual(stats_df["term"].tolist(), ["apple", "banana", "pear"])
        self.assertEqual(stats_df["count"].tolist(), [3, 3, 2])
        self.assertAlmostEqual(float(stats_df["share"].sum()), 1.0)

    def test_export_term_stats_writes_excel_output(self):
        stats_df = build_term_stats(["alpha beta alpha"], WordCloudConfig())

        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = f"{tmpdir}/terms.xlsx"
            saved_path = export_term_stats(stats_df, out_path)
            loaded = pd.read_excel(saved_path, engine="openpyxl")

        self.assertEqual(saved_path, out_path)
        self.assertEqual(loaded["term"].tolist(), ["alpha", "beta"])
        self.assertEqual(loaded["count"].tolist(), [2, 1])

    def test_render_wordcloud_requires_dependency_or_returns_image(self):
        stats_df = build_term_stats(["alpha beta alpha"], WordCloudConfig())
        config = WordCloudConfig(width=320, height=200)

        if wordcloud_tool.WordCloud is None:
            with self.assertRaises(ImportError):
                render_wordcloud(stats_df, config)
            return

        image = render_wordcloud(stats_df, config)
        self.assertEqual(image.size, (320, 200))


if __name__ == "__main__":
    unittest.main()
