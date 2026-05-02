"""Standard message-box helpers."""

from PySide6 import QtWidgets


def show_error(parent, title, message):
    QtWidgets.QMessageBox.critical(parent, title, str(message))


def show_warning(parent, title, message):
    QtWidgets.QMessageBox.warning(parent, title, str(message))


def show_info(parent, title, message):
    QtWidgets.QMessageBox.information(parent, title, str(message))
