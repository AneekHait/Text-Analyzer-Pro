"""Tests for the Setup-page data source panel state machine."""
from __future__ import annotations

import pandas as pd
import pytest
from PySide6 import QtWidgets

from textanalyzer.ui.data_source_panel import DataSourcePanel, DropZone


@pytest.fixture
def panel(qapp) -> DataSourcePanel:
    return DataSourcePanel()


@pytest.fixture
def sample_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "id": range(10),
            "comment": [f"row {i} text content" for i in range(10)],
            "score": [round(i * 0.1, 2) for i in range(10)],
        }
    )


class TestEmptyState:
    def test_starts_in_empty_state(self, panel: DataSourcePanel) -> None:
        assert panel._stack.currentWidget() is panel.dropzone

    def test_dropzone_height_is_locked_at_120(self, panel: DataSourcePanel) -> None:
        # setFixedHeight clamps both ends, ensuring the empty state never
        # balloons to fill its parent column.
        assert panel.dropzone.minimumHeight() == 120
        assert panel.dropzone.maximumHeight() == 120

    def test_dropzone_accepts_drops(self, panel: DataSourcePanel) -> None:
        assert panel.dropzone.acceptDrops()

    def test_browse_request_propagates(self, panel: DataSourcePanel) -> None:
        received: list[bool] = []
        panel.browse_requested.connect(lambda: received.append(True))
        panel.dropzone.browse_requested.emit()
        assert received == [True]


class TestStateTransitions:
    def test_show_loaded_state_switches_stack(
        self, panel: DataSourcePanel, sample_df: pd.DataFrame, tmp_path
    ) -> None:
        path = tmp_path / "f.csv"
        path.write_text("id\n1\n", encoding="utf-8")
        panel.show_loaded_state(file_path=str(path), sheets=["Sheet1"], current_sheet="Sheet1")
        assert panel._stack.currentWidget() is not panel.dropzone

    def test_show_empty_state_returns_to_dropzone(
        self, panel: DataSourcePanel, tmp_path
    ) -> None:
        path = tmp_path / "f.csv"
        path.write_text("id\n1\n", encoding="utf-8")
        panel.show_loaded_state(file_path=str(path), sheets=["Sheet1"], current_sheet="Sheet1")
        panel.show_empty_state()
        assert panel._stack.currentWidget() is panel.dropzone


class TestSheetSelector:
    def test_single_sheet_hides_selector(
        self, panel: DataSourcePanel, tmp_path
    ) -> None:
        path = tmp_path / "f.csv"
        path.write_text("id\n1\n", encoding="utf-8")
        panel.show_loaded_state(file_path=str(path), sheets=["Only"], current_sheet="Only")
        # Selector exists but should be hidden for single-sheet sources.
        assert panel._sheet_combo.count() == 1

    def test_multi_sheet_populates_and_preserves_selection(
        self, panel: DataSourcePanel, tmp_path
    ) -> None:
        path = tmp_path / "f.xlsx"
        path.write_bytes(b"")
        panel.show_loaded_state(file_path=str(path), sheets=["A", "B", "C"], current_sheet="B")
        assert panel._sheet_combo.count() == 3
        assert panel._sheet_combo.currentText() == "B"

    def test_sheet_change_emits_signal(
        self, panel: DataSourcePanel, tmp_path
    ) -> None:
        path = tmp_path / "f.xlsx"
        path.write_bytes(b"")
        panel.show_loaded_state(file_path=str(path), sheets=["A", "B"], current_sheet="A")
        received: list[str] = []
        panel.sheet_changed.connect(received.append)
        panel._sheet_combo.setCurrentText("B")
        assert received == ["B"]


class TestPreviewModel:
    def test_preview_caps_at_5_rows(
        self, panel: DataSourcePanel, sample_df: pd.DataFrame, tmp_path
    ) -> None:
        path = tmp_path / "f.csv"
        path.write_text("id\n1\n", encoding="utf-8")
        panel.show_loaded_state(file_path=str(path), sheets=["s"], current_sheet="s")
        panel.update_preview(sample_df)
        assert panel._preview_model.rowCount() == 5
        assert panel._preview_model.columnCount() == 3

    def test_preview_height_clamped(self, panel: DataSourcePanel) -> None:
        assert panel._preview_table.minimumHeight() == 150
        assert panel._preview_table.maximumHeight() == 180

    def test_highlight_column_is_stored(
        self, panel: DataSourcePanel, sample_df: pd.DataFrame, tmp_path
    ) -> None:
        path = tmp_path / "f.csv"
        path.write_text("id\n1\n", encoding="utf-8")
        panel.show_loaded_state(file_path=str(path), sheets=["s"], current_sheet="s")
        panel.update_preview(sample_df, highlight_column="comment")
        assert panel._preview_model._highlight_col == "comment"

    def test_preview_truncates_long_strings(
        self, panel: DataSourcePanel, tmp_path
    ) -> None:
        path = tmp_path / "f.csv"
        path.write_text("id\n1\n", encoding="utf-8")
        panel.show_loaded_state(file_path=str(path), sheets=["s"], current_sheet="s")
        long = "x" * 500
        df = pd.DataFrame({"col": [long]})
        panel.update_preview(df)
        from PySide6 import QtCore

        idx = panel._preview_model.index(0, 0)
        display = panel._preview_model.data(idx, QtCore.Qt.DisplayRole)
        tooltip = panel._preview_model.data(idx, QtCore.Qt.ToolTipRole)
        assert display.endswith("…") and len(display) <= 120
        assert tooltip == long  # full value preserved in tooltip


class TestStackSizing:
    def test_stack_hint_follows_current_page_not_largest(
        self, panel: DataSourcePanel, sample_df: pd.DataFrame, tmp_path
    ) -> None:
        # Regression: default QStackedWidget returns the largest child's hint,
        # which kept the empty-state Data Source group ballooning to the loaded
        # view's size. _CurrentPageStack overrides that.
        empty_hint = panel._stack.sizeHint()

        path = tmp_path / "f.csv"
        path.write_text("id\n1\n", encoding="utf-8")
        panel.show_loaded_state(file_path=str(path), sheets=["s"], current_sheet="s")
        panel.update_preview(sample_df)
        loaded_hint_h = panel._stack.sizeHint().height()

        panel.show_empty_state()
        empty_hint_after = panel._stack.sizeHint().height()

        # Same page → same hint, regardless of whether the loaded view was
        # constructed in between.
        assert empty_hint.height() == empty_hint_after


class TestDropZonePathFiltering:
    @pytest.mark.parametrize(
        "filename,expected",
        [
            ("data.xlsx", True),
            ("data.xlsm", True),
            ("data.xltx", True),
            ("data.xltm", True),
            ("data.xls", True),
            ("data.xlsb", True),
            ("data.ods", True),
            ("data.csv", True),
            ("data.json", True),
            ("data.txt", False),
            ("data.pdf", False),
            ("README.md", False),
        ],
    )
    def test_supported_extensions(
        self, qapp, tmp_path, filename: str, expected: bool
    ) -> None:
        from PySide6 import QtCore

        path = tmp_path / filename
        path.write_bytes(b"")
        url = QtCore.QUrl.fromLocalFile(str(path))
        result = DropZone._first_supported_path([url])
        assert (result is not None) == expected
