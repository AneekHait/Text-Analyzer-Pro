"""Smoke test for the main window's Setup wiring."""
from __future__ import annotations

import pytest


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
