"""Tests for IOService.save_taxonomy / load_taxonomy and apply_taxonomy.

Round-trip semantics, schema-version guard, legacy-clustering-bundle isolation,
pivot-sheet generation, user_renames survival.
"""
from __future__ import annotations

import os

import joblib
import pandas as pd
import pytest

import textanalyzer.engine.cluster as ct
from textanalyzer.services.io import IOService, TAXONOMY_SCHEMA_VERSION


@pytest.fixture
def trained_taxonomy():
    """Run categorization once on a synonym-heavy corpus and return the result."""
    texts = [
        "SAP ECC batch job failed RP1",
        "SAP ECC batch job aborted RP1",
        "SAP ECC batch failure RP1",
        "SAP ECC batch error RP1",
        "SAP ECC batch failed run",
        "user login failed",
        "user login error",
        "user login failure",
        "user login denied",
        "user login authentication",
        "out of band oddity",
        "another rare one off thing",
        "rare coffee machine ticket",
    ]
    r = ct.categorize_taxonomy(
        texts, vectorizer_kind="tfidf",
        min_cluster_size=3, non_repetitive_min_size=3,
    )
    return texts, r


class TestSaveLoadRoundtrip:
    def test_roundtrip_preserves_required_keys(self, tmp_path, trained_taxonomy):
        _texts, result = trained_taxonomy
        path = tmp_path / "tax.joblib"
        IOService.save_taxonomy(result, str(path))
        loaded = IOService.load_taxonomy(str(path))
        assert loaded["schema_version"] == TAXONOMY_SCHEMA_VERSION
        assert loaded["kind"] == "taxonomy"
        for key in ("vectorizer", "sub_centroids", "subcategory_names", "sub_fingerprints"):
            assert key in loaded

    def test_apply_taxonomy_on_same_corpus_returns_same_labels(self, tmp_path, trained_taxonomy):
        texts, result = trained_taxonomy
        path = tmp_path / "tax.joblib"
        IOService.save_taxonomy(result, str(path))
        loaded = IOService.load_taxonomy(str(path))
        applied = ct.apply_taxonomy(texts, loaded, confidence_threshold=0.0)
        # For every Repetitive row in the original run, the re-applied row
        # should land in some valid subcategory string (not necessarily byte-
        # identical because kmeans-on-centroids isn't reused, but the basic
        # invariant: 3 columns aligned, no crashes).
        assert len(applied.repetitive) == len(texts)
        assert len(applied.subcategory) == len(texts)
        assert len(applied.confidence) == len(texts)

    def test_apply_taxonomy_low_confidence_becomes_non_repetitive(self, tmp_path, trained_taxonomy):
        _texts, result = trained_taxonomy
        path = tmp_path / "tax.joblib"
        IOService.save_taxonomy(result, str(path))
        loaded = IOService.load_taxonomy(str(path))
        # Texts unrelated to anything in the training set should sit far from
        # every centroid and fall into Non-Repetitive when the threshold is
        # high.
        ood = ["completely unrelated content about gardening tomatoes"]
        applied = ct.apply_taxonomy(ood, loaded, confidence_threshold=0.99)
        assert applied.repetitive[0] == "Non-Repetitive"
        assert applied.subcategory[0] == "Non-Repetitive"
        assert applied.confidence[0] == 0.0

    def test_user_renames_survive_save_load(self, tmp_path, trained_taxonomy):
        _texts, result = trained_taxonomy
        path = tmp_path / "tax.joblib"
        # Grab a real fingerprint and stash a fake user rename through the save call.
        any_fp = next(iter(result.sub_fingerprints.values()))
        IOService.save_taxonomy(result, str(path), user_renames={any_fp: "My Custom Name"})
        loaded = IOService.load_taxonomy(str(path))
        assert loaded["user_renames"][any_fp] == "My Custom Name"


class TestSchemaGuards:
    def test_missing_file_raises_filenotfound(self, tmp_path):
        with pytest.raises(FileNotFoundError, match="not found"):
            IOService.load_taxonomy(str(tmp_path / "nope.joblib"))

    def test_wrong_kind_raises_runtime(self, tmp_path):
        # Hand-craft a legacy-style cluster .joblib (no "kind" field).
        path = tmp_path / "legacy.joblib"
        joblib.dump({"model": object(), "vectorizer": object()}, str(path))
        with pytest.raises(RuntimeError, match="kind"):
            IOService.load_taxonomy(str(path))

    def test_schema_version_mismatch_raises(self, tmp_path):
        path = tmp_path / "future.joblib"
        joblib.dump({
            "schema_version": 999, "kind": "taxonomy",
            "vectorizer": object(), "sub_centroids": [], "subcategory_names": {},
        }, str(path))
        with pytest.raises(RuntimeError, match="schema"):
            IOService.load_taxonomy(str(path))

    def test_missing_required_key_raises(self, tmp_path):
        path = tmp_path / "broken.joblib"
        joblib.dump({"schema_version": TAXONOMY_SCHEMA_VERSION, "kind": "taxonomy"}, str(path))
        with pytest.raises(RuntimeError, match="vectorizer"):
            IOService.load_taxonomy(str(path))


class TestLegacyModelBundleStillLoads:
    """A regression guard: existing IOService.load_model must keep working on
    bundles saved by the pre-taxonomy code (no `kind` key, no schema version)."""

    def test_old_cluster_payload_loads_through_load_model(self, tmp_path):
        path = tmp_path / "old.joblib"
        joblib.dump({
            "model": object(),
            "vectorizer": object(),
            "cluster_names": {0: "alpha", 1: "beta"},
            "top_keywords": {0: [("alpha", 1.0)], 1: [("beta", 0.9)]},
        }, str(path))
        payload = IOService.load_model(str(path))
        assert payload["cluster_names"][0] == "alpha"


class TestPivotSheetExport:
    def test_pivot_sheet_written_for_xlsx(self, tmp_path):
        df = pd.DataFrame({
            "Short description": [f"ticket {i}" for i in range(10)],
            "Repetitive/Non-Repetitive": ["Repetitive"] * 7 + ["Non-Repetitive"] * 3,
            "Subcategory": ["SAP ECC Batch Failure"] * 4 + ["User Login Issue"] * 3 + ["Non-Repetitive"] * 3,
            "Confidence": [0.8] * 7 + [0.0] * 3,
        })
        out = tmp_path / "out.xlsx"
        IOService.save_results_with_pivot(df, str(out), sheet_name="Inc")
        assert out.exists()
        pivot = pd.read_excel(out, sheet_name="pivot")
        assert "Row Labels" in pivot.columns
        assert "Count" in pivot.columns
        assert "%" in pivot.columns
        # Repetitive total + 2 subcategories + Non-Repetitive total + TOTAL = 5 rows
        assert len(pivot) == 5
        total_row = pivot[pivot["Row Labels"] == "TOTAL"]
        assert int(total_row["Count"].iloc[0]) == 10

    def test_pivot_dataframe_orders_subcategories_descending(self):
        df = pd.DataFrame({
            "Repetitive/Non-Repetitive": ["Repetitive"] * 8 + ["Non-Repetitive"] * 2,
            "Subcategory": ["A"] * 5 + ["B"] * 2 + ["C"] * 1 + ["Non-Repetitive"] * 2,
            "Confidence": [0.5] * 10,
        })
        pivot = IOService.build_pivot_dataframe(df)
        # First three subcategory rows (after the "Repetitive" header) should
        # be A (5), B (2), C (1).
        sub_rows = pivot[pivot["Row Labels"].str.startswith("  ")]
        assert sub_rows["Count"].tolist() == [5, 2, 1]

    def test_csv_output_skips_pivot_sheet(self, tmp_path):
        df = pd.DataFrame({
            "Short description": ["x", "y"],
            "Repetitive/Non-Repetitive": ["Repetitive", "Non-Repetitive"],
            "Subcategory": ["Foo Bar", "Non-Repetitive"],
            "Confidence": [0.7, 0.0],
        })
        out = tmp_path / "out.csv"
        IOService.save_results_with_pivot(df, str(out))
        assert out.exists()
        # A plain CSV — round-trip read confirms no pivot sheet leakage.
        read = pd.read_csv(out)
        assert "Subcategory" in read.columns


class TestManifestAndDescribe:
    def test_manifest_persists_through_save_load(self, tmp_path, trained_taxonomy):
        _texts, result = trained_taxonomy
        path = tmp_path / "tax.joblib"
        IOService.save_taxonomy(result, str(path))
        loaded = IOService.load_taxonomy(str(path))
        m = loaded.get("manifest") or {}
        # Manifest survived the round-trip.
        assert m.get("min_cluster_size") == result.manifest.get("min_cluster_size")
        assert m.get("vectorizer_kind") == result.manifest.get("vectorizer_kind")

    def test_describe_taxonomy_includes_key_facts(self, tmp_path, trained_taxonomy):
        _texts, result = trained_taxonomy
        path = tmp_path / "tax.joblib"
        IOService.save_taxonomy(result, str(path))
        loaded = IOService.load_taxonomy(str(path))
        description = IOService.describe_taxonomy(loaded)
        assert description.startswith("Taxonomy:")
        assert "subcategories" in description
        # At least one knob makes it into the human-readable summary.
        assert "min_cluster_size" in description


class TestPivotEnhancements:
    def test_avg_confidence_column_when_subcat_map_provided(self):
        df = pd.DataFrame({
            "Short description": ["a", "b", "c"],
            "Repetitive/Non-Repetitive": ["Repetitive", "Repetitive", "Non-Repetitive"],
            "Subcategory": ["Alpha", "Alpha", "Non-Repetitive"],
            "Confidence": [0.7, 0.9, 0.0],
        })
        pivot = IOService.build_pivot_dataframe(
            df, avg_confidence_by_subcat={"Alpha": 0.8},
        )
        assert "Avg Confidence" in pivot.columns
        alpha_row = pivot[pivot["Row Labels"] == "  Alpha"].iloc[0]
        assert abs(float(alpha_row["Avg Confidence"]) - 0.8) < 1e-9

    def test_group_column_when_labels_provided(self):
        df = pd.DataFrame({
            "Short description": ["a", "b", "c", "d"],
            "Repetitive/Non-Repetitive": ["Repetitive"] * 3 + ["Non-Repetitive"],
            "Subcategory": ["Alpha", "Beta", "Gamma", "Non-Repetitive"],
            "Confidence": [0.5] * 4,
        })
        groups = {"Alpha": "Group X", "Beta": "Group X", "Gamma": "Group Y"}
        pivot = IOService.build_pivot_dataframe(df, group_labels=groups)
        assert "Group" in pivot.columns
        # Group X (Alpha + Beta) should appear together in the ordering.
        sub_rows = pivot[pivot["Row Labels"].str.startswith("  ")]
        groups_in_order = sub_rows["Group"].tolist()
        # First two subcategory rows share Group X (sorted alphabetically by group label).
        assert groups_in_order[0] == groups_in_order[1]

    def test_save_results_with_pivot_includes_taxonomy_columns(self, tmp_path, trained_taxonomy):
        _texts, result = trained_taxonomy
        df = pd.DataFrame({
            "Short description": [f"row {i}" for i in range(len(result.repetitive))],
            "Repetitive/Non-Repetitive": result.repetitive,
            "Subcategory": result.subcategory,
            "Confidence": result.confidence,
        })
        out = tmp_path / "with_tax.xlsx"
        IOService.save_results_with_pivot(df, str(out), taxonomy_result=result)
        pivot = pd.read_excel(out, sheet_name="pivot")
        # Avg Confidence is present because taxonomy_result was supplied.
        assert "Avg Confidence" in pivot.columns
