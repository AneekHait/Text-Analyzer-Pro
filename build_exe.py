"""Build script for creating TextAnalyzerPro.exe via PyInstaller."""

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).parent

def main():
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", "TextAnalyzerPro",
        "--windowed",
        "--icon", str(ROOT / "assets" / "logo.ico"),
        "--add-data", f"{ROOT / 'assets'};assets",
        "--add-data", f"{ROOT / 'ascii_banner.txt'};.",
        "--add-data", f"{ROOT / 'theme'};theme",
        "--add-data", f"{ROOT / 'textanalyzer'};textanalyzer",
        "--hidden-import", "PySide6.QtSvg",
        "--hidden-import", "sklearn.cluster",
        "--hidden-import", "sklearn.feature_extraction.text",
        "--hidden-import", "sklearn.decomposition",
        "--hidden-import", "sklearn.manifold",
        "--hidden-import", "sklearn.metrics",
        "--hidden-import", "sklearn.neighbors",
        "--hidden-import", "hdbscan",
        "--hidden-import", "openpyxl",
        "--hidden-import", "xlrd",
        "--hidden-import", "odfpy",
        "--hidden-import", "wordcloud",
        "--hidden-import", "qtawesome",
        "--noconfirm",
        "--clean",
        str(ROOT / "gui.py"),
    ]
    print("Running PyInstaller...")
    print(" ".join(cmd))
    result = subprocess.run(cmd)
    if result.returncode == 0:
        print(f"\nBuild successful! EXE at: dist/TextAnalyzerPro/TextAnalyzerPro.exe")
    else:
        print(f"\nBuild failed with code {result.returncode}")
    return result.returncode

if __name__ == "__main__":
    sys.exit(main())
