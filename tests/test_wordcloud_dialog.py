"""Tests for the new column selector and live-render plumbing in WordCloudDialog."""
from __future__ import annotations

import pandas as pd
import pytest


@pytest.fixture
def dialog_factory(qapp, monkeypatch):
    """Build a WordCloudDialog whose _generate_preview is patched to a no-op.

    Mirrors the trick used in test_wordcloud_tool.py: the real renderer needs
    the wordcloud package and is slow + not what we want to exercise here.
    """
    from textanalyzer.ui.wordcloud_window import WordCloudDialog

    def make(texts=None, column_name="text", dataframe=None):
        monkeypatch.setattr(WordCloudDialog, "_generate_preview", lambda self: None)
        return WordCloudDialog(None, texts or ["alpha beta", "gamma"], column_name, dataframe=dataframe)

    return make


class TestColumnSelector:
    def test_source_group_hidden_without_dataframe(self, dialog_factory) -> None:
        d = dialog_factory()
        # When no dataframe is supplied, the Source group is hidden — combo
        # exists (cleaner than conditional attributes) but its parent is not
        # visible.
        assert d.column_combo.parent().isVisibleTo(d) is False or d.column_combo.count() == 0

    def test_combo_populates_from_dataframe(self, dialog_factory) -> None:
        df = pd.DataFrame({"text": ["a", "b"], "id": [1, 2], "comment": ["x", "y"]})
        d = dialog_factory(column_name="comment", dataframe=df)
        items = [d.column_combo.itemText(i) for i in range(d.column_combo.count())]
        assert "text" in items and "comment" in items and "id" in items

    def test_string_columns_listed_first(self, dialog_factory) -> None:
        df = pd.DataFrame({"id": [1, 2], "comment": ["x", "y"], "score": [0.1, 0.2]})
        d = dialog_factory(column_name="comment", dataframe=df)
        items = [d.column_combo.itemText(i) for i in range(d.column_combo.count())]
        # 'comment' is object/string → should come before numeric 'id' / 'score'.
        assert items.index("comment") < items.index("id")
        assert items.index("comment") < items.index("score")

    def test_initial_selection_matches_column_name(self, dialog_factory) -> None:
        df = pd.DataFrame({"a": ["x"], "b": ["y"], "c": ["z"]})
        d = dialog_factory(column_name="b", dataframe=df)
        assert d.column_combo.currentText() == "b"

    def test_column_change_updates_texts_and_title(self, dialog_factory) -> None:
        df = pd.DataFrame({"left": ["one two", "three"], "right": ["alpha", "beta"]})
        d = dialog_factory(column_name="left", dataframe=df)
        d._on_column_change("right")
        assert d.column_name == "right"
        assert d.texts == ["alpha", "beta"]
        assert "right" in d.windowTitle()

    def test_column_change_to_nonexistent_is_noop(self, dialog_factory) -> None:
        df = pd.DataFrame({"a": ["x"]})
        d = dialog_factory(column_name="a", dataframe=df)
        original_texts = list(d.texts)
        d._on_column_change("does-not-exist")
        assert d.texts == original_texts


class TestLiveRender:
    def test_debounce_timer_is_single_shot(self, dialog_factory) -> None:
        d = dialog_factory()
        assert d._debounce_timer.isSingleShot()
        assert d._debounce_timer.interval() >= 100  # not so tight it floods

    def test_schedule_regen_starts_timer(self, dialog_factory) -> None:
        d = dialog_factory()
        assert not d._debounce_timer.isActive()
        d._schedule_regen()
        assert d._debounce_timer.isActive()

    def test_settings_change_triggers_regen(self, dialog_factory) -> None:
        d = dialog_factory()
        d._debounce_timer.stop()
        # Pick a value different from the default so currentTextChanged fires.
        d.color_combo.setCurrentText("Rainbow Mix")
        assert d._debounce_timer.isActive()

        d._debounce_timer.stop()
        d.max_words_spin.setValue(d.max_words_spin.value() + 10)
        assert d._debounce_timer.isActive()

        d._debounce_timer.stop()
        d.rel_scale_slider.setValue(50)
        assert d._debounce_timer.isActive()

    def test_column_change_triggers_regen(self, dialog_factory) -> None:
        df = pd.DataFrame({"a": ["one"], "b": ["two"]})
        d = dialog_factory(column_name="a", dataframe=df)
        d._debounce_timer.stop()
        d._on_column_change("b")
        assert d._debounce_timer.isActive()
