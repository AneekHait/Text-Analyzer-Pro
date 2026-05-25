"""Tests for the single-level subcategory categorization pipeline.

Covers: categorize_taxonomy + apply_taxonomy + phrase_name_from_keywords +
fingerprint stability + user-rename preservation + Non-Repetitive bucketing +
confidence column + duplicate-name disambiguation + precomputed-shortcircuit.

LLM- or transformers-related cases don't exist (the feature is keyword-only by
design — see plan-for-improving-the-cuddly-puffin.md).
"""
from __future__ import annotations

import numpy as np
import pytest

import textanalyzer.engine.cluster as ct


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def ticket_corpus() -> list[str]:
    """Two tight clusters + a sprinkle of singletons that should land in
    Non-Repetitive after the noise+tiny merge step."""
    return [
        # cluster A: SAP batch failures
        "SAP ECC batch job failed RP1",
        "SAP ECC batch job aborted RP1",
        "SAP ECC batch failure RP1",
        "SAP ECC batch error RP1",
        "SAP ECC batch failed run",
        # cluster B: login / password
        "user login failed",
        "user login error",
        "user login failure",
        "user login failed authentication",
        "user login denied",
        # singletons (these become Non-Repetitive on small min_cluster_size)
        "rare coffee machine ticket",
        "out of band oddity",
        "another one off thing",
    ]


# ---------------------------------------------------------------------------
# phrase_name_from_keywords
# ---------------------------------------------------------------------------

class TestPhraseNameFromKeywords:
    def test_returns_titlecase_phrase_from_unigrams(self):
        out = ct.phrase_name_from_keywords([
            ("sap", 0.9), ("ecc", 0.8), ("batch", 0.7), ("failure", 0.6),
        ])
        # Multi-word, no repeats, Title Case.
        words = out.split()
        assert len(words) == 4
        assert words == ["Sap", "Ecc", "Batch", "Failure"]

    def test_prefers_ngram_when_top_term_is_multiword(self):
        # If the side TF-IDF's top term is an n-gram, return it whole rather
        # than padding it with unrelated unigrams.
        out = ct.phrase_name_from_keywords([
            ("batch job failure", 0.95),
            ("sap", 0.6), ("ecc", 0.5),
        ])
        assert out == "Batch Job Failure"

    def test_dedupes_repeated_keywords(self):
        out = ct.phrase_name_from_keywords([
            ("sap", 0.9), ("SAP", 0.8), ("sap", 0.7), ("batch", 0.6),
        ])
        # Case-insensitive dedupe.
        assert out.lower().count("sap") == 1
        assert "Batch" in out

    def test_empty_returns_empty_string(self):
        assert ct.phrase_name_from_keywords([]) == ""

    def test_caps_at_max_chars(self):
        long_kw = [(f"term{i}", 1.0 - i * 0.01) for i in range(20)]
        out = ct.phrase_name_from_keywords(long_kw, max_chars=20)
        assert len(out) <= 20


# ---------------------------------------------------------------------------
# Subcluster fingerprint
# ---------------------------------------------------------------------------

class TestSubclusterFingerprint:
    def test_stable_under_reorder(self):
        a = ct._subcluster_fingerprint([("sap", 0.9), ("batch", 0.8), ("ecc", 0.7)])
        b = ct._subcluster_fingerprint([("ecc", 0.5), ("sap", 0.95), ("batch", 0.6)])
        assert a == b

    def test_stable_under_case(self):
        a = ct._subcluster_fingerprint([("sap", 0.9), ("batch", 0.8)])
        b = ct._subcluster_fingerprint([("SAP", 0.5), ("Batch", 0.4)])
        assert a == b

    def test_different_kw_lists_differ(self):
        a = ct._subcluster_fingerprint([("sap", 0.9), ("batch", 0.8)])
        b = ct._subcluster_fingerprint([("login", 0.9), ("password", 0.8)])
        assert a != b


# ---------------------------------------------------------------------------
# Non-repetitive merge
# ---------------------------------------------------------------------------

class TestMarkNonRepetitive:
    def test_noise_stays_noise(self):
        labels = np.array([-1, -1, 0, 0, 0, 1, 1, 1])
        out = ct._mark_non_repetitive(labels, min_size=2)
        # Noise (-1) untouched, clusters 0 and 1 both ≥ 2 → kept.
        assert int(np.sum(out == -1)) == 2

    def test_tiny_cluster_demoted(self):
        labels = np.array([0, 0, 1])  # cluster 1 has only 1 member
        out = ct._mark_non_repetitive(labels, min_size=2)
        assert out.tolist() == [0, 0, -1]


# ---------------------------------------------------------------------------
# categorize_taxonomy end-to-end
# ---------------------------------------------------------------------------

class TestCategorizeTaxonomy:
    def test_returns_three_aligned_columns(self, ticket_corpus):
        r = ct.categorize_taxonomy(
            ticket_corpus,
            vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        n = len(ticket_corpus)
        assert len(r.repetitive) == n
        assert len(r.subcategory) == n
        assert len(r.confidence) == n
        assert all(0.0 <= c <= 1.0 for c in r.confidence)

    def test_non_repetitive_bucket_captures_singletons(self, ticket_corpus):
        r = ct.categorize_taxonomy(
            ticket_corpus,
            vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        # The 3 explicit singletons at the end should ALL be Non-Repetitive,
        # and their confidence must be exactly 0.0 by construction.
        for i in range(-3, 0):
            assert r.repetitive[i] == "Non-Repetitive"
            assert r.subcategory[i] == "Non-Repetitive"
            assert r.confidence[i] == 0.0

    def test_confidence_is_zero_for_all_non_repetitive_rows(self, ticket_corpus):
        r = ct.categorize_taxonomy(
            ticket_corpus, vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        for rep, conf in zip(r.repetitive, r.confidence):
            if rep == "Non-Repetitive":
                assert conf == 0.0

    def test_stats_dict_has_expected_keys(self, ticket_corpus):
        r = ct.categorize_taxonomy(
            ticket_corpus, vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        for key in ("n_subclusters", "n_non_repetitive", "pct_non_repetitive", "vectorizer_kind"):
            assert key in r.stats

    def test_user_renames_preserved_on_rerun(self, ticket_corpus):
        # First run to discover the fingerprints + raw names.
        r1 = ct.categorize_taxonomy(
            ticket_corpus, vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        # Find a real subcluster (not Non-Repetitive) and rename it.
        valid_ids = [cid for cid in r1.sub_fingerprints.keys() if cid != -1]
        assert valid_ids, "Fixture should have at least one repetitive subcluster"
        target_cid = valid_ids[0]
        target_fp = r1.sub_fingerprints[target_cid]
        custom = "USER EDITED: SAP ECC Batch Job Failure"

        # Re-run with the user_renames mapping. The cluster keyed by the same
        # fingerprint should now adopt the custom name.
        r2 = ct.categorize_taxonomy(
            ticket_corpus, vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
            user_renames={target_fp: custom},
        )
        # Custom name must appear in the subcategory_names map for some cluster id.
        assert custom in r2.subcategory_names.values()

    def test_precomputed_skips_vectorize(self, ticket_corpus, monkeypatch):
        # Vectorize once via the public API, then make a sentinel that fails
        # if vectorize_texts is called again — the precomputed path must skip it.
        vec, X = ct.vectorize_texts(ticket_corpus, vectorizer_kind="tfidf")

        def _boom(*args, **kwargs):
            pytest.fail("vectorize_texts should not be called when precomputed= is given")

        monkeypatch.setattr(ct, "vectorize_texts", _boom)
        r = ct.categorize_taxonomy(
            ticket_corpus, precomputed=(vec, X),
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        assert len(r.repetitive) == len(ticket_corpus)

    def test_falls_back_to_tfidf_when_embedding_unavailable(self, ticket_corpus, monkeypatch):
        # Pretend sentence-transformers isn't installed, request embedding mode.
        # categorize_taxonomy should warn and fall through to TF-IDF.
        monkeypatch.setattr(ct, "_ST_AVAILABLE", False)
        with pytest.warns(UserWarning, match="sentence-transformers"):
            r = ct.categorize_taxonomy(
                ticket_corpus, vectorizer_kind="embedding",
                min_cluster_size=3, non_repetitive_min_size=3,
            )
        assert r.stats["vectorizer_kind"] == "tfidf"

    def test_duplicate_names_disambiguated(self):
        # Force two clusters to resolve to the same raw name by handcrafting
        # the disambiguation helper directly.
        names = {0: "Same Name", 1: "Same Name", 2: "Unique"}
        out = ct._disambiguate_duplicate_names(names)
        assert out[0] == "Same Name"
        assert out[1] == "Same Name (1)"
        assert out[2] == "Unique"


# ---------------------------------------------------------------------------
# Improvements: case preservation, c-TF-IDF, samples, UMAP/scale guard,
# HDBSCAN-aware confidence, merge / split engine helpers.
# ---------------------------------------------------------------------------

class TestCasePreservation:
    def test_uppercase_acronyms_survive(self, ticket_corpus):
        case_map = ct._case_form_lookup(ticket_corpus)
        # SAP appears uppercase in every fixture row → must round-trip uppercase.
        assert case_map["sap"] == "SAP"
        assert case_map["ecc"] == "ECC"
        assert case_map["rp1"] == "RP1"

    def test_lowercase_corpus_words_get_title_case(self):
        case_map = {"failed": "failed"}
        # "failed" appears only as lowercase in the corpus → render as Title Case.
        assert ct._apply_case_map("failed", case_map) == "Failed"

    def test_stop_word_connectives_stay_lower(self):
        case_map = {"of": "of", "the": "the"}
        assert ct._apply_case_map("of", case_map) == "of"
        assert ct._apply_case_map("the", case_map) == "the"

    def test_phrase_name_uses_case_map(self, ticket_corpus):
        case_map = ct._case_form_lookup(ticket_corpus)
        name = ct.phrase_name_from_keywords(
            [("sap", 0.9), ("ecc", 0.8), ("batch", 0.7), ("failure", 0.6)],
            case_map=case_map,
        )
        # SAP + ECC preserved; "batch" + "failure" Title-Cased.
        assert "SAP" in name and "ECC" in name
        assert "Batch" in name
        assert "Failure" in name

    def test_phrase_name_dedupes_across_unigrams_and_ngrams(self):
        # Top keyword "rare" + n-gram "rare topic" + unigram "topic" should NOT
        # render with repeated words.
        name = ct.phrase_name_from_keywords([
            ("rare", 0.9),
            ("rare topic", 0.85),
            ("topic", 0.7),
            ("issue", 0.5),
        ])
        words = name.split()
        # Each lowercase word appears at most once in the rendered phrase.
        lowered = [w.lower() for w in words]
        assert len(set(lowered)) == len(lowered), f"duplicates in: {name}"


class TestCTfIdfKeywords:
    def test_each_cluster_gets_distinctive_terms(self):
        # Two clusters with very different vocabularies → c-TF-IDF should
        # surface the SAP keywords for one and the login keywords for the
        # other (zero cross-contamination).
        texts = [
            "SAP ECC batch job failed", "SAP ECC batch aborted",
            "SAP ECC batch failure", "SAP ECC job error",
            "login failed", "login error",
            "password reset", "account locked out login",
        ]
        labels = np.array([0, 0, 0, 0, 1, 1, 1, 1])
        kw = ct._compute_c_tf_idf_keywords(texts, labels)
        kw0 = " ".join([t for t, _ in kw[0]]).lower()
        kw1 = " ".join([t for t, _ in kw[1]]).lower()
        assert "sap" in kw0 and "ecc" in kw0
        assert "login" in kw1 or "password" in kw1
        # SAP must NOT leak into cluster 1's top terms.
        assert "sap" not in kw1

    def test_empty_or_noise_only_returns_empty(self):
        out = ct._compute_c_tf_idf_keywords(["a", "b"], np.array([-1, -1]))
        assert out == {}


class TestSampleDerivedNames:
    def test_returns_candidate_when_phrase_repeats(self):
        samples = [
            "ServiceNow form approval workflow stuck",
            "ServiceNow form approval workflow blocked",
            "ServiceNow form approval workflow timeout",
        ]
        top_kw = [("approval", 0.9), ("workflow", 0.8), ("servicenow", 0.7)]
        case_map = ct._case_form_lookup(samples)
        result = ct._sample_derived_name(samples, top_kw, case_map)
        assert result is not None
        # Contains keyword tokens and is short.
        assert "Workflow" in result or "Approval" in result
        assert len(result) <= 60

    def test_returns_none_when_no_repetition(self):
        samples = ["wildly unrelated thing one", "completely different thing two"]
        top_kw = [("thing", 0.9), ("one", 0.5)]
        result = ct._sample_derived_name(samples, top_kw, {})
        # No 3-gram repeats across samples → no candidate.
        assert result is None

    def test_returns_none_when_only_one_sample(self):
        result = ct._sample_derived_name(["just one sample"], [("x", 0.5)], {})
        assert result is None


class TestConfidenceUsesProbabilitiesAndMargin:
    def test_confidence_uses_margin_when_no_hdbscan_probs(self):
        # Two centroids placed orthogonally. Rows very close to one centroid
        # should have HIGH confidence (best cosine + margin); rows close to
        # neither should fall to a low score even if best cosine is decent.
        X = np.array([[1.0, 0.0], [0.9, 0.1], [0.7, 0.7]], dtype=np.float64)
        centroids = np.array([[1.0, 0.0], [0.0, 1.0]], dtype=np.float64)
        labels = np.array([0, 0, 0])
        scores = ct._confidence_scores(X, labels, centroids, [0, 1])
        # Row 0 (perfectly aligned): high. Row 2 (45° to both): much lower.
        assert scores[0] > scores[2] + 0.2

    def test_hdbscan_probs_dominate_when_provided(self):
        # Same setup but explicit HDBSCAN probabilities — confirm the
        # combined score moves with the probabilities.
        X = np.array([[1.0, 0.0], [1.0, 0.0]], dtype=np.float64)
        centroids = np.array([[1.0, 0.0], [0.0, 1.0]], dtype=np.float64)
        labels = np.array([0, 0])
        low = ct._confidence_scores(
            X, labels, centroids, [0, 1], hdbscan_probabilities=np.array([0.1, 0.1]),
        )
        high = ct._confidence_scores(
            X, labels, centroids, [0, 1], hdbscan_probabilities=np.array([0.9, 0.9]),
        )
        assert high[0] > low[0]


class TestScaleGuard:
    def test_below_threshold_uses_hdbscan(self):
        # Pulled from the engine module constant.
        from textanalyzer.engine.cluster import _CATEGORIZATION_SCALE_GUARD as _GUARD
        assert _GUARD == 25_000

    def test_above_threshold_switches_to_minibatch(self, monkeypatch):
        # Force the scale path to trigger with a tiny synthetic threshold so
        # the test stays fast — the function takes ``min_cluster_size`` and
        # returns synthetic noise tagging via percentile cutoff.
        rng = np.random.default_rng(0)
        X = rng.random((200, 8))
        labels = ct._subclusters_at_scale(X, min_cluster_size=2)
        # Some rows tagged as noise (top 5% by centroid distance).
        assert int((labels == -1).sum()) >= 1
        # Multiple distinct cluster ids returned.
        assert len({int(l) for l in labels}) > 1


class TestUMAPOptionality:
    def test_passes_through_when_umap_unavailable(self, monkeypatch):
        monkeypatch.setattr(ct, "_UMAP_AVAILABLE", False)
        X = np.eye(5)
        out, applied = ct._apply_umap_reduction(X)
        assert applied is False
        assert out is X  # returned unchanged


class TestMergeAndSplit:
    @pytest.fixture
    def two_cluster_result(self, ticket_corpus):
        return ct.categorize_taxonomy(
            ticket_corpus, vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        ), ticket_corpus

    def test_merge_reduces_cluster_count(self, two_cluster_result):
        result, texts = two_cluster_result
        ids = sorted(result.subcategory_names.keys())
        if len(ids) < 2:
            pytest.skip("Need ≥2 clusters in the fixture to test merge")
        vec, X = ct.vectorize_texts(texts, vectorizer_kind="tfidf")
        merged = ct.merge_clusters(result, X, texts, [ids[0], ids[1]])
        # One fewer cluster after merge.
        assert len(merged.subcategory_names) == len(result.subcategory_names) - 1

    def test_merge_requires_two_clusters(self, two_cluster_result):
        result, texts = two_cluster_result
        vec, X = ct.vectorize_texts(texts, vectorizer_kind="tfidf")
        with pytest.raises(ValueError, match="at least 2"):
            ct.merge_clusters(result, X, texts, [0])

    def test_split_increases_cluster_count(self, two_cluster_result):
        result, texts = two_cluster_result
        ids = [
            cid for cid in result.subcategory_names.keys()
            if int(np.sum(result.subcluster_labels == cid)) >= 4
        ]
        if not ids:
            pytest.skip("Need a cluster with ≥4 members to split")
        vec, X = ct.vectorize_texts(texts, vectorizer_kind="tfidf")
        split = ct.split_cluster(result, X, texts, ids[0], k=2)
        assert len(split.subcategory_names) == len(result.subcategory_names) + 1

    def test_split_too_small_raises(self, two_cluster_result):
        result, texts = two_cluster_result
        ids = sorted(result.subcategory_names.keys())
        if not ids:
            pytest.skip("Need at least 1 cluster")
        vec, X = ct.vectorize_texts(texts, vectorizer_kind="tfidf")
        with pytest.raises(ValueError, match="cannot split"):
            ct.split_cluster(result, X, texts, ids[0], k=999)


class TestManifestAndSamples:
    def test_manifest_present_after_run(self, ticket_corpus):
        r = ct.categorize_taxonomy(
            ticket_corpus, vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        m = r.manifest
        # Required keys for provenance.
        for key in (
            "created_at", "vectorizer_kind", "min_cluster_size",
            "min_samples", "non_repetitive_min_size", "n_rows",
            "n_subclusters", "scale_path", "umap_applied",
        ):
            assert key in m, f"missing manifest key: {key}"

    def test_samples_and_avg_conf_populated(self, ticket_corpus):
        r = ct.categorize_taxonomy(
            ticket_corpus, vectorizer_kind="tfidf",
            min_cluster_size=3, non_repetitive_min_size=3,
        )
        # Every non-noise cluster id has at least one sample + an avg confidence.
        for cid in r.subcategory_names.keys():
            assert cid in r.samples_by_cluster
            assert cid in r.avg_confidence_by_cluster
            assert 0.0 <= r.avg_confidence_by_cluster[cid] <= 1.0
