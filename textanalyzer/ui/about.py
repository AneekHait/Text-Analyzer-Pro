"""About dialog with brand banner and quick links."""

import webbrowser

from PySide6 import QtCore, QtWidgets

from ..utils.icons import _icon, _load_app_icon, _load_banner_pixmap


class AboutDialog(QtWidgets.QDialog):
    def __init__(self, owner_name, owner_contact, owner_website, owner_bmc, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Text Analyzer Pro")
        self.setMinimumWidth(520)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(28, 22, 28, 18)
        layout.setSpacing(10)

        # Header banner
        banner_pixmap = _load_banner_pixmap(140)
        if not banner_pixmap.isNull():
            banner_label = QtWidgets.QLabel()
            banner_label.setPixmap(banner_pixmap)
            banner_label.setAlignment(QtCore.Qt.AlignCenter)
            layout.addWidget(banner_label)
            version = QtWidgets.QLabel("Version 1.5  •  MIT License")
            version.setProperty("role", "muted")
            version.setAlignment(QtCore.Qt.AlignCenter)
            layout.addWidget(version)
        else:
            header = QtWidgets.QHBoxLayout()
            logo_label = QtWidgets.QLabel()
            app_icon = _load_app_icon()
            if not app_icon.isNull():
                logo_label.setPixmap(app_icon.pixmap(72, 72))
            header.addWidget(logo_label)
            text_col = QtWidgets.QVBoxLayout()
            title = QtWidgets.QLabel("Text Analyzer Pro")
            title.setProperty("role", "h1")
            version = QtWidgets.QLabel("Version 1.5  •  MIT License")
            version.setProperty("role", "muted")
            text_col.addWidget(title)
            text_col.addWidget(version)
            text_col.addStretch(1)
            header.addLayout(text_col, 1)
            layout.addLayout(header)

        body = QtWidgets.QLabel(
            "<p style='line-height:1.45; margin:0'>"
            "Intelligent text clustering and word-cloud analysis for Excel, CSV, and JSON files. "
            "All processing happens locally on your machine — nothing is uploaded.</p>"
            f"<p style='margin:8px 0 0 0'><b>Owner:</b> {owner_name}</p>"
        )
        body.setTextFormat(QtCore.Qt.RichText)
        body.setWordWrap(True)
        body.setOpenExternalLinks(True)
        layout.addWidget(body)

        layout.addSpacing(12)

        # Link buttons
        button_row = QtWidgets.QHBoxLayout()
        link_specs = [
            ("LinkedIn", owner_contact, "fa5b.linkedin"),
            ("Website", owner_website, "fa5s.globe"),
            ("Buy Me a Coffee", owner_bmc, "fa5s.mug-hot"),
        ]
        for label, url, icon_name in link_specs:
            if not url:
                continue
            btn = QtWidgets.QPushButton(label)
            btn.setProperty("flat", "true")
            btn.setIcon(_icon(icon_name))
            btn.clicked.connect(lambda _checked=False, target=url: webbrowser.open(target))
            button_row.addWidget(btn)
        button_row.addStretch(1)
        close_button = QtWidgets.QPushButton("Close")
        close_button.setProperty("primary", "true")
        close_button.setDefault(True)
        close_button.setAutoDefault(True)
        close_button.clicked.connect(self.accept)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)

        close_button.setFocus()
