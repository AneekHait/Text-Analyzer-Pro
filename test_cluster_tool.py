import os
import tempfile
import unittest
from unittest import mock

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pandas as pd
from PySide6 import QtWidgets

from cluster_tool import (
    TextCleaningConfig,
    get_sheet_names,
    load_table,
    prepare_text_cleaning,
    preprocess_texts,
    save_results,
)
from gui import ClusterGUI


def get_qapp():
    app = QtWidgets.QApplication.instance()
    if app is None:
        app = QtWidgets.QApplication([])
    return app


class TextCleaningTests(unittest.TestCase):
    def test_preprocess_texts_default_behavior_is_compatible(self):
        result = preprocess_texts(["  Hello  ", "WORLD", None])
        self.assertEqual(result, ["hello", "world", ""])

    def test_text_cleaning_operations_apply_together(self):
        config = TextCleaningConfig(
            remove_urls=True,
            remove_emails=True,
            remove_numbers=True,
            remove_punctuation=True,
            regex_pattern=r"alpha",
            regex_replacement="omega",
        )
        result = preprocess_texts([" Alpha! 123 test@example.com https://example.com "], config)
        self.assertEqual(result, ["omega"])

    def test_invalid_regex_is_rejected(self):
        with self.assertRaises(ValueError):
            TextCleaningConfig(regex_pattern="(")

    def test_prepare_text_cleaning_reports_dedupe_and_empty_rows(self):
        config = TextCleaningConfig(
            dedupe_cleaned_rows=True,
            remove_numbers=True,
            collapse_whitespace=True,
            trim_whitespace=True,
            lowercase=True,
        )
        result = prepare_text_cleaning(["Alpha 123", " alpha ", None, "Beta"], config)
        self.assertEqual(result.cleaned_texts, ["alpha", "alpha", "", "beta"])
        self.assertEqual(result.cluster_input_texts, ["alpha", "beta"])
        self.assertEqual(result.kept_indices, [0, 3])
        self.assertEqual(result.representative_index_by_row, [0, 0, None, 3])
        self.assertEqual(result.stats["deduped_row_count"], 1)
        self.assertEqual(result.stats["empty_row_count"], 1)


class FileFormatSupportTests(unittest.TestCase):
    def test_load_table_reads_csv_and_json(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = os.path.join(tmpdir, "demo.csv")
            json_path = os.path.join(tmpdir, "demo.json")
            source = pd.DataFrame({"text": ["alpha", "beta"], "value": [1, 2]})
            source.to_csv(csv_path, index=False)
            source.to_json(json_path, orient="records", indent=2)
            csv_df = load_table(csv_path)
            json_df = load_table(json_path)
        self.assertEqual(csv_df.to_dict(orient="records"), source.to_dict(orient="records"))
        self.assertEqual(json_df.to_dict(orient="records"), source.to_dict(orient="records"))

    def test_get_sheet_names_returns_single_table_for_csv_and_json(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = os.path.join(tmpdir, "demo.csv")
            json_path = os.path.join(tmpdir, "demo.json")
            pd.DataFrame({"text": ["alpha"]}).to_csv(csv_path, index=False)
            pd.DataFrame({"text": ["alpha"]}).to_json(json_path, orient="records")
            self.assertEqual(get_sheet_names(csv_path), ["Data"])
            self.assertEqual(get_sheet_names(json_path), ["Data"])

    def test_save_results_writes_csv_and_json(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = os.path.join(tmpdir, "clustered.csv")
            json_path = os.path.join(tmpdir, "clustered.json")
            source = pd.DataFrame({"text": ["alpha"], "cluster_label": [0], "cluster_name": ["topic"]})
            saved_csv_path = save_results(source, csv_path)
            saved_json_path = save_results(source, json_path)
            loaded_csv = pd.read_csv(saved_csv_path)
            loaded_json = pd.read_json(saved_json_path)
        self.assertEqual(loaded_csv.to_dict(orient="records"), source.to_dict(orient="records"))
        self.assertEqual(loaded_json.to_dict(orient="records"), source.to_dict(orient="records"))


class ClusterGuiCleaningTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.qapp = get_qapp()

    def setUp(self):
        self.gui = ClusterGUI()
        self.gui.df = pd.DataFrame({"text": [" Alpha 123 ", "alpha", None, "Beta!!!"]})
        self.gui.current_file_path = os.path.join(os.getcwd(), "demo.xlsx")
        self.gui.col_combo.addItem("text")
        self.gui.col_combo.setCurrentText("text")
        self.gui.refresh_cleaning_preview()

    def tearDown(self):
        self.gui.close()

    def test_cleaning_preview_updates_summary(self):
        self.gui.remove_numbers_check.setChecked(True)
        self.gui.refresh_cleaning_preview()
        self.assertEqual(self.gui.cleaning_source_label.text(), "text")
        self.assertEqual(self.gui.cleaning_output_label.text(), "text_cleaned")
        self.assertEqual(self.gui.cleaning_rows_label.text(), "4")
        self.assertEqual(self.gui.cleaning_preview_table.rowCount(), 4)

    def test_session_recipe_load_restores_form(self):
        self.gui.remove_urls_check.setChecked(True)
        self.gui.remove_numbers_check.setChecked(True)
        self.gui.cleaning_recipes["Demo"] = self.gui._cleaning_config_to_dict()
        self.gui._refresh_cleaning_recipe_combo()
        self.gui.remove_urls_check.setChecked(False)
        self.gui.remove_numbers_check.setChecked(False)
        self.gui.cleaning_recipe_combo.setCurrentText("Demo")
        self.gui.load_selected_cleaning_recipe()
        self.assertTrue(self.gui.remove_urls_check.isChecked())
        self.assertTrue(self.gui.remove_numbers_check.isChecked())

    def test_run_clustering_creates_cleaned_column_and_maps_duplicate_labels(self):
        self.gui.remove_numbers_check.setChecked(True)
        self.gui.remove_punctuation_check.setChecked(True)
        self.gui.dedupe_cleaned_rows_check.setChecked(True)
        self.gui.k_spin.setValue(2)
        with mock.patch("gui.show_error"), mock.patch("gui.show_warning"):
            self.gui.run_clustering()
        self.assertIn("text_cleaned", self.gui.df.columns)
        self.assertEqual(self.gui.df["text_cleaned"].tolist(), ["alpha", "alpha", "", "beta"])
        self.assertEqual(len(self.gui.labels), 4)
        self.assertEqual(int(self.gui.labels[0]), int(self.gui.labels[1]))
        self.assertEqual(int(self.gui.labels[2]), -1)


if __name__ == "__main__":
    unittest.main()
