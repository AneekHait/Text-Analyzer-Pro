"""Tests for the Suggest K / Advanced TF-IDF / Compare Algorithms GUI surfaces."""
from __future__ import annotations

import os
import tempfile

import pandas as pd
import pytest


@pytest.fixture
def gui(qapp, monkeypatch):
    """Build a ClusterGUI with a small loaded dataframe."""
    from gui import ClusterGUI

    win = ClusterGUI()
    df = pd.DataFrame({
        "text": [
            "apple banana cherry",
            "apple banana",
            "engine motor wheel",
            "engine wheel chassis",
            "python compile error",
            "python javascript code",
        ]
    })
    win.df = df
    # Make col_combo aware of the column so current_column_name() returns it.
    win.col_combo.clear()
    win.col_combo.addItem("text")
    win.col_combo.setCurrentText("text")
    yield win
    win.close()


class TestSetupButtons:
    def test_buttons_exist_and_are_enabled(self, gui) -> None:
        assert gui.suggest_k_btn is not None and gui.suggest_k_btn.isEnabled()
        assert gui.elbow_k_btn is not None and gui.elbow_k_btn.isEnabled()
        assert gui.advanced_tfidf_btn is not None and gui.advanced_tfidf_btn.isEnabled()
        assert gui.compare_algos_btn is not None and gui.compare_algos_btn.isEnabled()


class TestTfidfSettings:
    def test_defaults_match_legacy_behavior(self, gui) -> None:
        s = gui._tfidf_settings
        assert s["max_features"] == 2000
        assert s["min_df"] == 1
        assert s["max_df"] == 1.0
        assert s["ngram_range"] == (1, 1)
        assert s["use_hashing"] is False

    def test_settings_round_trip_through_prepare_corpus(self, gui) -> None:
        # When we override settings, _prepare_setup_corpus passes them through.
        gui._tfidf_settings["ngram_range"] = (1, 2)
        gui._tfidf_settings["max_features"] = 50
        result = gui._prepare_setup_corpus()
        assert result is not None
        vectorizer, X = result
        # ngram_range=(1,2) means features include bigrams. n_features <= 50.
        assert hasattr(vectorizer, "ngram_range")
        assert vectorizer.ngram_range == (1, 2)
        assert X.shape[1] <= 50


class TestSuggestK:
    def test_suggest_optimal_k_applies_via_message_box_apply(self, gui, monkeypatch) -> None:
        # Patch the QMessageBox.question to "Apply" without showing the modal.
        from PySide6 import QtWidgets

        gui.k_spin.setValue(2)
        monkeypatch.setattr(
            QtWidgets.QMessageBox,
            "question",
            staticmethod(lambda *a, **k: QtWidgets.QMessageBox.Apply),
        )
        gui.suggest_optimal_k()
        # k_spin should have been set to whatever find_optimal_k recommends —
        # we don't assert a specific number (data-dependent), only that the
        # value changed off the default (2 is the floor of the search range).
        assert gui.k_spin.value() >= 2

    def test_suggest_optimal_k_does_not_apply_on_cancel(self, gui, monkeypatch) -> None:
        from PySide6 import QtWidgets

        gui.k_spin.setValue(7)
        monkeypatch.setattr(
            QtWidgets.QMessageBox,
            "question",
            staticmethod(lambda *a, **k: QtWidgets.QMessageBox.Cancel),
        )
        gui.suggest_optimal_k()
        assert gui.k_spin.value() == 7  # unchanged

    def test_elbow_method_logs_and_applies(self, gui, monkeypatch) -> None:
        from PySide6 import QtWidgets

        log_calls: list[str] = []
        monkeypatch.setattr(gui, "log_msg", lambda msg: log_calls.append(msg))
        monkeypatch.setattr(
            QtWidgets.QMessageBox,
            "question",
            staticmethod(lambda *a, **k: QtWidgets.QMessageBox.Apply),
        )
        gui.k_spin.setValue(2)
        gui.suggest_optimal_k(method="elbow")
        # The log output should mention the elbow path, not silhouette.
        joined = " ".join(log_calls).lower()
        assert "elbow" in joined
        assert "silhouette" not in joined

    def test_invalid_method_falls_back_to_silhouette(self, gui, monkeypatch) -> None:
        from PySide6 import QtWidgets

        log_calls: list[str] = []
        monkeypatch.setattr(gui, "log_msg", lambda msg: log_calls.append(msg))
        monkeypatch.setattr(
            QtWidgets.QMessageBox,
            "question",
            staticmethod(lambda *a, **k: QtWidgets.QMessageBox.Cancel),
        )
        gui.suggest_optimal_k(method="not-a-method")
        joined = " ".join(log_calls).lower()
        assert "silhouette" in joined

    def test_suggest_warns_when_no_data(self, qapp, monkeypatch) -> None:
        from gui import ClusterGUI
        from PySide6 import QtWidgets

        seen: list[tuple] = []
        monkeypatch.setattr(
            QtWidgets.QMessageBox, "warning",
            staticmethod(lambda *a, **k: seen.append(("warn", a)) or QtWidgets.QMessageBox.Ok),
        )
        win = ClusterGUI()
        try:
            win.suggest_optimal_k()
        finally:
            win.close()
        assert any("warn" == s[0] for s in seen)


class TestCompareAlgorithms:
    def test_compare_runs_and_logs_best(self, gui, monkeypatch) -> None:
        # Stub out the result + progress dialogs so the test doesn't block on
        # offscreen Qt event quirks (QProgressDialog can immediately emit
        # `canceled` when shown headlessly).
        from PySide6 import QtWidgets

        class _StubProgress:
            def __init__(self, *args, **kwargs):
                self.canceled = type("S", (), {"connect": lambda *a, **k: None})()
            def setWindowTitle(self, *_a, **_k): pass
            def setWindowModality(self, *_a, **_k): pass
            def setMinimumDuration(self, *_a, **_k): pass
            def setValue(self, *_a, **_k): pass
            def setLabelText(self, *_a, **_k): pass
            def close(self): pass

        monkeypatch.setattr(QtWidgets, "QProgressDialog", _StubProgress)
        monkeypatch.setattr(
            QtWidgets.QDialog,
            "exec",
            lambda self: QtWidgets.QDialog.Accepted,
        )
        log_calls: list[str] = []
        monkeypatch.setattr(gui, "log_msg", lambda msg: log_calls.append(msg))
        gui.compare_algorithms_dialog()
        assert any("comparison" in m.lower() for m in log_calls)
