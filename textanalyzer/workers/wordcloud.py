"""QObject-based wordcloud rendering worker.

Runs :func:`prepare_wordcloud_data` + :func:`render_wordcloud` on a worker
thread so the UI remains responsive. A monotonically-increasing
``request_id`` lets the UI ignore stale results when the user changes
inputs while a render is in flight.
"""

from __future__ import annotations

from PySide6 import QtCore

from textanalyzer.engine.wordcloud import prepare_wordcloud_data, render_wordcloud


class WordCloudWorker(QtCore.QObject):
    finished = QtCore.Signal(int, str, object, object, object)  # request_id, column, stats_df, summary, image
    empty = QtCore.Signal(int, str, object)  # request_id, column, summary
    failed = QtCore.Signal(int, str)  # request_id, error message

    def __init__(self, request_id: int, column: str, texts: list[str], config) -> None:
        super().__init__()
        self.request_id = request_id
        self.column = column
        self.texts = texts
        self.config = config

    @QtCore.Slot()
    def run(self) -> None:
        try:
            stats_df, summary = prepare_wordcloud_data(self.texts, self.config)
            if stats_df is None or stats_df.empty:
                self.empty.emit(self.request_id, self.column, summary)
                return
            image = render_wordcloud(stats_df, self.config)
            self.finished.emit(self.request_id, self.column, stats_df, summary, image)
        except Exception as error:
            self.failed.emit(self.request_id, str(error))


__all__ = ["WordCloudWorker"]
