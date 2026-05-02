"""Worker that runs cleaning preview on a background thread.

Replaces the synchronous ``prepare_text_cleaning`` call in
``refresh_cleaning_preview`` so the UI stays responsive during large files.
"""

from __future__ import annotations

from PySide6 import QtCore

from textanalyzer.engine.cluster import prepare_text_cleaning, TextCleaningConfig


class CleaningPreviewWorker(QtCore.QObject):
    """Run ``prepare_text_cleaning`` off the main thread."""

    finished = QtCore.Signal(object)   # TextCleaningResult
    failed = QtCore.Signal(str)

    def __init__(self, texts: list[str], config: TextCleaningConfig) -> None:
        super().__init__()
        self.texts = texts
        self.config = config

    @QtCore.Slot()
    def run(self) -> None:
        try:
            result = prepare_text_cleaning(self.texts, self.config)
            self.finished.emit(result)
        except Exception as exc:
            self.failed.emit(str(exc))


__all__ = ["CleaningPreviewWorker"]
