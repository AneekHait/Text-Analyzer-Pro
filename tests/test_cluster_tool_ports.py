"""Tests for the features ported back from the distributable into cluster_tool.

Covers: validate_input, find_optimal_k, compare_algorithms, ApplicableModel,
HashingVectorizer path in vectorize_texts, MiniBatchKMeans switching, and
HDBSCAN routing.
"""
from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

import cluster_tool as ct


@pytest.fixture
def clean_corpus() -> list[str]:
    """A small but well-separated 4-cluster corpus."""
    return [
        # cluster 1: animals
        "the cat sat on the mat",
        "dog and cat play together",
        "puppy chased the kitten",
        # cluster 2: cars
        "the engine of the car broke",
        "diesel engines and gasoline engines differ",
        "tire pressure affects engine fuel",
        # cluster 3: software
        "python code compiled cleanly",
        "javascript and python both have functions",
        "compile error in the source code",
        # cluster 4: cooking
        "chop the onions and add salt",
        "saute garlic with butter and pepper",
        "salt and pepper for seasoning the soup",
    ]


@pytest.fixture
def df_for_validate() -> pd.DataFrame:
    return pd.DataFrame({
        "text": [
            "alpha beta gamma",
            "alpha beta",
            "delta epsilon zeta",
            "delta epsilon",
            "",
            None,
            "  ",
            "iota kappa lambda",
        ],
        "id": list(range(8)),
    })


# ---------------------------------------------------------------------------
# validate_input
# ---------------------------------------------------------------------------

class TestValidateInput:
    def test_unknown_column_returns_invalid(self, df_for_validate):
        result = ct.validate_input(df_for_validate, "nope", algorithm="kmeans", n_clusters=2)
        assert result.is_valid is False
        assert any("not found" in e.lower() for e in result.errors)

    def test_stats_include_null_and_empty(self, df_for_validate):
        result = ct.validate_input(df_for_validate, "text", algorithm="kmeans", n_clusters=2)
        s = result.stats
        assert s["total_rows"] == 8
        assert s["null_count"] == 1  # the None
        # Empty count includes "", "  " (after strip), and None (filled with "")
        assert s["empty_count"] >= 3
        assert "avg_text_length" in s
        assert "unique_texts" in s

    def test_n_clusters_too_high_warns(self, df_for_validate):
        result = ct.validate_input(df_for_validate, "text", algorithm="kmeans", n_clusters=20)
        # 20 > 8 non-empty rows → error
        assert result.is_valid is False
        assert any("greater than" in e.lower() for e in result.errors)

    def test_dbscan_carries_noise_warning(self, df_for_validate):
        result = ct.validate_input(df_for_validate, "text", algorithm="dbscan")
        assert any("noise" in w.lower() for w in result.warnings)

    def test_all_empty_column_is_invalid(self):
        df = pd.DataFrame({"col": ["", "  ", None, ""]})
        result = ct.validate_input(df, "col", algorithm="kmeans", n_clusters=2)
        assert result.is_valid is False


# ---------------------------------------------------------------------------
# vectorize_texts — TF-IDF knobs and HashingVectorizer
# ---------------------------------------------------------------------------

class TestVectorizer:
    def test_default_returns_tfidf(self, clean_corpus):
        vec, X = ct.vectorize_texts(clean_corpus, max_features=200)
        from sklearn.feature_extraction.text import TfidfVectorizer
        assert isinstance(vec, TfidfVectorizer)
        assert X.shape[0] == len(clean_corpus)

    def test_ngram_range_increases_features(self, clean_corpus):
        _, X1 = ct.vectorize_texts(clean_corpus, max_features=500, ngram_range=(1, 1))
        _, X2 = ct.vectorize_texts(clean_corpus, max_features=500, ngram_range=(1, 2))
        assert X2.shape[1] > X1.shape[1]

    def test_hashing_path_returns_hashing_vectorizer(self, clean_corpus):
        from sklearn.feature_extraction.text import HashingVectorizer
        vec, X = ct.vectorize_texts(clean_corpus, max_features=64, use_hashing=True, chunk_size=4)
        assert isinstance(vec, HashingVectorizer)
        assert X.shape == (len(clean_corpus), 64)

    def test_hashing_top_keywords_returns_empty(self, clean_corpus):
        # HashingVectorizer has no feature names — get_top_keywords_per_cluster
        # should warn and return empty per-cluster keyword lists rather than crashing.
        vec, X = ct.vectorize_texts(clean_corpus, max_features=64, use_hashing=True)
        labels = np.array([0, 0, 1, 1, 1, 0, 1, 0, 0, 1, 0, 1])
        with pytest.warns(UserWarning):
            kw = ct.get_top_keywords_per_cluster(vec, X, labels, top_n=3)
        assert all(len(v) == 0 for v in kw.values())


# ---------------------------------------------------------------------------
# find_optimal_k
# ---------------------------------------------------------------------------

class TestFindOptimalK:
    def test_silhouette_picks_a_k_in_range(self, clean_corpus):
        _, X = ct.vectorize_texts(clean_corpus, max_features=100)
        result = ct.find_optimal_k(X, k_range=(2, 6), method="silhouette")
        assert 2 <= result["optimal_k"] <= 6
        assert result["confidence"] in {"high", "medium", "low", "very_low"}
        assert "Recommended" in result["recommendation"]

    def test_elbow_method_returns_a_k(self, clean_corpus):
        _, X = ct.vectorize_texts(clean_corpus, max_features=100)
        result = ct.find_optimal_k(X, k_range=(2, 6), method="elbow")
        assert result["optimal_k"] >= 2
        assert "elbow" in result["recommendation"].lower()


# ---------------------------------------------------------------------------
# compare_algorithms
# ---------------------------------------------------------------------------

class TestCompareAlgorithms:
    def test_returns_one_result_per_algorithm(self, clean_corpus):
        _, X = ct.vectorize_texts(clean_corpus, max_features=100)
        results = ct.compare_algorithms(X, n_clusters=4, eps=0.3, min_samples=2)
        names = [r.name for r in results]
        assert names == ["kmeans", "dbscan", "agglomerative"]
        # KMeans should always produce a valid silhouette on this corpus.
        kmeans_result = next(r for r in results if r.name == "kmeans")
        assert kmeans_result.silhouette is not None

    def test_get_best_algorithm_picks_max_silhouette(self):
        results = [
            ct.AlgorithmResult("kmeans", np.array([]), 3, 0.4, None, None, 0, 0.0),
            ct.AlgorithmResult("dbscan", np.array([]), 2, 0.6, None, None, 0, 0.0),
            ct.AlgorithmResult("agglo", np.array([]), 3, 0.5, None, None, 0, 0.0),
        ]
        assert ct.get_best_algorithm(results) == "dbscan"

    def test_get_best_algorithm_returns_none_when_no_silhouette(self):
        results = [ct.AlgorithmResult("kmeans", np.array([]), 0, None, None, None, 0, 0.0)]
        assert ct.get_best_algorithm(results) is None


# ---------------------------------------------------------------------------
# ApplicableModel
# ---------------------------------------------------------------------------

class TestApplicableModel:
    def test_kmeans_delegates_to_native_predict(self, clean_corpus):
        _, X = ct.vectorize_texts(clean_corpus, max_features=100)
        model, labels = ct.cluster_texts(X, algorithm="kmeans", n_clusters=3)
        wrapped = ct.wrap_model_for_prediction(model, X, labels, "kmeans")
        # On the same data, predictions should equal the training labels exactly.
        pred = wrapped.predict(X)
        assert np.array_equal(pred, labels)

    def test_dbscan_uses_nearest_neighbor(self, clean_corpus):
        _, X = ct.vectorize_texts(clean_corpus, max_features=100)
        model, labels = ct.cluster_texts(X, algorithm="dbscan", eps=0.5, min_samples=2)
        wrapped = ct.wrap_model_for_prediction(model, X, labels, "dbscan")
        # Each training row's nearest neighbor is itself → predictions match labels.
        pred = wrapped.predict(X)
        assert np.array_equal(pred, labels)

    def test_agglomerative_uses_nearest_neighbor(self, clean_corpus):
        _, X = ct.vectorize_texts(clean_corpus, max_features=100)
        model, labels = ct.cluster_texts(X, algorithm="agglomerative", n_clusters=3)
        wrapped = ct.wrap_model_for_prediction(model, X, labels, "agglomerative")
        pred = wrapped.predict(X)
        assert np.array_equal(pred, labels)


# ---------------------------------------------------------------------------
# cluster_texts — MiniBatchKMeans switch + HDBSCAN routing
# ---------------------------------------------------------------------------

class TestClusterRouting:
    def test_minibatch_kicks_in_above_threshold(self):
        # Synthetic 50-row dataset, threshold=10 → MiniBatchKMeans should be chosen.
        rng = np.random.default_rng(0)
        X = rng.random((50, 8))
        model, labels = ct.cluster_texts(
            X, algorithm="kmeans", n_clusters=3, minibatch_threshold=10
        )
        from sklearn.cluster import MiniBatchKMeans
        assert isinstance(model, MiniBatchKMeans)
        assert len(labels) == 50

    def test_minibatch_does_not_kick_in_below_threshold(self):
        rng = np.random.default_rng(0)
        X = rng.random((20, 8))
        model, _ = ct.cluster_texts(X, algorithm="kmeans", n_clusters=3, minibatch_threshold=100)
        from sklearn.cluster import KMeans, MiniBatchKMeans
        assert isinstance(model, KMeans) and not isinstance(model, MiniBatchKMeans)

    def test_unknown_algorithm_raises(self):
        rng = np.random.default_rng(0)
        X = rng.random((10, 4))
        with pytest.raises(ValueError):
            ct.cluster_texts(X, algorithm="not-a-real-algo", n_clusters=2)

    def test_hdbscan_raises_clear_error_when_unavailable(self):
        # If hdbscan isn't installed, calling with algorithm='hdbscan' should
        # raise ImportError with an actionable hint. Otherwise it should work.
        rng = np.random.default_rng(0)
        X = rng.random((30, 8))
        if ct._HDBSCAN_AVAILABLE:
            model, labels = ct.cluster_texts(X, algorithm="hdbscan", min_cluster_size=3)
            assert len(labels) == 30
        else:
            with pytest.raises(ImportError, match="hdbscan"):
                ct.cluster_texts(X, algorithm="hdbscan")


# ---------------------------------------------------------------------------
# TextCleaningConfig — lemmatize / custom_stopwords
# ---------------------------------------------------------------------------

class TestCleaningConfigExtensions:
    def test_custom_stopwords_filtered(self):
        cfg = ct.TextCleaningConfig(custom_stopwords=("the", "and"))
        out = ct.clean_text_value("The cat AND the dog", cfg)
        assert "the" not in out.split()
        assert "and" not in out.split()
        assert "cat" in out.split()
        assert "dog" in out.split()

    def test_lemmatize_flag_does_not_crash_when_unavailable(self):
        # Even if NLTK isn't installed or corpora aren't present, this should
        # return a string (best-effort), not raise.
        cfg = ct.TextCleaningConfig(lemmatize=True)
        out = ct.clean_text_value("The cats are running quickly", cfg)
        assert isinstance(out, str) and len(out) > 0


# ---------------------------------------------------------------------------
# load_table — encoding fallback chain
# ---------------------------------------------------------------------------

class TestLoadTableEncodings:
    def test_load_table_reads_cp1252_csv(self, tmp_path):
        """Excel-style CSVs with cp1252 bytes (e.g. en-dash 0x96) load cleanly."""
        path = tmp_path / "windows.csv"
        # 0x96 is en-dash in cp1252 — would raise UnicodeDecodeError under utf-8.
        # Build the bytes directly: cp1252 has no mapping for Python's U+0096
        # codepoint, so we splice the raw byte in.
        path.write_bytes(b"name,note\nAlice,en\x96dash\nBob,plain\n")
        with pytest.warns(UserWarning, match="cp1252"):
            df = ct.load_table(str(path))
        assert list(df.columns) == ["name", "note"]
        assert df.iloc[0]["note"] == "en–dash"  # cp1252 0x96 → U+2013 en-dash

    def test_load_table_reads_utf8_bom_csv(self, tmp_path):
        path = tmp_path / "bom.csv"
        # Leading BOM ﻿ would otherwise leak into the first column header.
        path.write_text("﻿name,note\nAlice,hello\n", encoding="utf-8")
        df = ct.load_table(str(path))
        assert list(df.columns) == ["name", "note"]
        assert df.iloc[0]["name"] == "Alice"

    def test_load_table_plain_utf8_csv_does_not_warn(self, tmp_path, recwarn):
        path = tmp_path / "clean.csv"
        path.write_text("a,b\n1,2\n", encoding="utf-8")
        df = ct.load_table(str(path))
        assert list(df.columns) == ["a", "b"]
        # No encoding-related warning should be emitted on a clean utf-8 file.
        encoding_warns = [w for w in recwarn.list if "is not UTF-8" in str(w.message)]
        assert encoding_warns == []

    def test_load_table_reads_cp1252_json(self, tmp_path):
        """Records-array JSON with cp1252 bytes loads via the same fallback."""
        path = tmp_path / "windows.json"
        path.write_bytes(b'[{"name": "Alice", "note": "en\x96dash"}]')
        with pytest.warns(UserWarning, match="cp1252"):
            df = ct.load_table(str(path))
        assert df.iloc[0]["note"] == "en–dash"


class TestLoadTableExcelFormats:
    """Verify pandas auto-picks the correct engine for each Excel-family format."""

    def test_xlsm_round_trip(self, tmp_path):
        # openpyxl handles .xlsm read+write the same as .xlsx (macros are
        # preserved in the binary blob but never executed by openpyxl).
        path = tmp_path / "macro.xlsm"
        pd.DataFrame({"a": [1, 2], "b": ["hello", "world"]}).to_excel(
            path, index=False, engine="openpyxl"
        )
        df = ct.load_table(str(path))
        assert list(df.columns) == ["a", "b"]
        assert df.iloc[0]["b"] == "hello"
        # get_sheet_names should also recognize the .xlsm extension.
        assert ct.get_sheet_names(str(path))  # at least one sheet

    def test_ods_round_trip(self, tmp_path):
        # odfpy is needed for both reading and writing .ods.
        odf = pytest.importorskip("odf")  # confirms odfpy is installed
        del odf  # imported only as availability check
        path = tmp_path / "calc.ods"
        pd.DataFrame({"x": [10, 20], "y": ["foo", "bar"]}).to_excel(
            path, index=False, engine="odf"
        )
        df = ct.load_table(str(path))
        assert list(df.columns) == ["x", "y"]
        assert df.iloc[1]["y"] == "bar"
        assert ct.get_sheet_names(str(path))

    def test_unsupported_extension_message_lists_all_excel_formats(self, tmp_path):
        path = tmp_path / "foo.parquet"
        path.write_bytes(b"")
        with pytest.raises(ValueError) as exc_info:
            ct.load_table(str(path))
        msg = str(exc_info.value)
        # Error message should mention every supported Excel extension.
        for ext in (".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"):
            assert ext in msg

    def test_excel_extension_set_includes_all_formats(self):
        assert ct.EXCEL_INPUT_EXTENSIONS == {
            ".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".ods"
        }
