"""Smoke test for the main window's Setup wiring."""
from __future__ import annotations

import pytest


@pytest.fixture(autouse=True)
def _isolated_settings(tmp_path, monkeypatch):
    """Redirect persistent settings to a tmpdir so tests don't clobber real prefs.

    The main window persists choices (last_vectorizer_kind, last_algorithm,
    recent_files, …) to ``~/.text_analyzer_pro/settings.json`` as the user
    clicks around. Tests must not write to the actual user file.
    """
    from textanalyzer import settings as app_settings

    monkeypatch.setattr(app_settings, "SETTINGS_DIR", tmp_path)
    monkeypatch.setattr(app_settings, "SETTINGS_FILE", tmp_path / "settings.json")
    # Suppress QSettings mirror as well — it would still hit the real registry.
    monkeypatch.setattr(app_settings, "_mirror_to_qsettings", lambda data: None)
    monkeypatch.setattr(app_settings, "_load_from_qsettings", lambda: {})
    yield


@pytest.fixture
def main_window(qapp):
    from gui import ClusterGUI

    win = ClusterGUI()
    yield win
    win.close()


def test_main_window_constructs(main_window) -> None:
    assert main_window is not None


def test_data_source_panel_attached(main_window) -> None:
    from textanalyzer.ui.data_source_panel import DataSourcePanel

    assert hasattr(main_window, "data_source_panel")
    assert isinstance(main_window.data_source_panel, DataSourcePanel)


def test_legacy_widget_shims_present(main_window) -> None:
    # The old `file_label` / `sheet_combo` attributes are still referenced
    # elsewhere in the gui module; the panel exposes them as compat shims.
    assert hasattr(main_window, "file_label")
    assert not main_window.file_label.isVisible()
    assert main_window.sheet_combo is main_window.data_source_panel._sheet_combo


def test_starts_in_empty_state(main_window) -> None:
    panel = main_window.data_source_panel
    assert panel._stack.currentWidget() is panel.dropzone


def test_vectorizer_combo_has_both_modes(main_window) -> None:
    """Setup tab exposes a vectorizer combo with TF-IDF (default) and Embeddings."""
    combo = main_window.vec_combo
    kinds = [combo.itemData(i) for i in range(combo.count())]
    assert kinds == ["tfidf", "embedding"]
    # Default is TF-IDF (matches both the engine default and a freshly-loaded settings file).
    assert main_window.current_vectorizer_kind() == "tfidf"


def test_build_vectorize_kwargs_switches_on_combo(main_window) -> None:
    """Flipping the combo to embeddings (when available) rewires vectorize_kwargs."""
    import textanalyzer.engine.cluster as ct
    kw_tfidf = main_window._build_vectorize_kwargs()
    assert kw_tfidf["vectorizer_kind"] == "tfidf"
    assert "max_features" in kw_tfidf

    if not ct._ST_AVAILABLE:
        pytest.skip("sentence-transformers not installed; embedding combo is disabled")
    main_window.vec_combo.setCurrentIndex(1)
    kw_embed = main_window._build_vectorize_kwargs()
    assert kw_embed["vectorizer_kind"] == "embedding"
    assert kw_embed["embedding_model"].startswith("sentence-transformers/")
    assert kw_embed["embedding_device"] in {"cpu", "cuda"}
    assert isinstance(kw_embed["embedding_batch_size"], int)


def test_embeddings_advanced_button_disabled_in_tfidf_mode(main_window) -> None:
    """The Embeddings… button is only enabled when the combo is in embedding mode."""
    assert main_window.advanced_embed_btn.isEnabled() is False
    assert main_window.advanced_tfidf_btn.isEnabled() is True


def test_run_categorization_button_present(main_window) -> None:
    """Action bar exposes the categorization button next to Run Clustering."""
    btn = main_window.categorize_btn
    assert btn is not None
    assert btn.isEnabled() is True
    assert btn.text() == "Run Categorization"


def test_save_taxonomy_action_starts_disabled(main_window) -> None:
    """Save Taxonomy is only available after a successful categorization run."""
    assert main_window.save_taxonomy_action.isEnabled() is False
    assert main_window.load_taxonomy_action.isEnabled() is True


def test_running_state_disables_categorize_button(main_window) -> None:
    """_set_running_state(True) disables BOTH Run Clustering and Run Categorization."""
    main_window._set_running_state(True)
    assert main_window.run_btn.isEnabled() is False
    assert main_window.categorize_btn.isEnabled() is False
    main_window._set_running_state(False)
    assert main_window.run_btn.isEnabled() is True
    assert main_window.categorize_btn.isEnabled() is True


def test_granularity_slider_formula_maps_distinct_values() -> None:
    """The documented mapping produces distinct mcs values at 0/50/100."""
    from textanalyzer.models.config import CategorizationConfig
    low = CategorizationConfig.min_cluster_size_from_granularity(0)
    mid = CategorizationConfig.min_cluster_size_from_granularity(50)
    high = CategorizationConfig.min_cluster_size_from_granularity(100)
    # Coarse (low g) → larger clusters; Fine (high g) → smaller min_cluster_size.
    assert low > mid > high
    assert high >= 3  # floor enforced


def test_controller_has_taxonomy_edit_signal(main_window) -> None:
    """The controller exposes edit_taxonomy so the merge/split handlers
    can adopt a rebuilt TaxonomyResult without a re-run."""
    assert hasattr(main_window.controller, "edit_taxonomy")
    assert callable(main_window.controller.edit_taxonomy)


def test_controller_initial_taxonomy_state(main_window) -> None:
    """Fresh window has empty taxonomy state until a categorization runs."""
    assert main_window.controller.taxonomy_result is None
    assert main_window.controller.taxonomy_X is None
    assert main_window.controller.taxonomy_texts == []
    assert main_window.controller.user_taxonomy_renames == {}
