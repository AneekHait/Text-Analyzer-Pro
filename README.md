# Text Analyzer Pro (v1.5)

Intelligent Text Clustering & Analysis — desktop GUI for clustering text from Excel workbooks.

Features
- Multi-sheet Excel support; choose any text column
- Preprocessing + TF‑IDF vectorization
- Clustering: K-Means, DBSCAN, Agglomerative
- Automatic keyword extraction and suggested human-readable cluster names
- 2D visualizations (PCA, t-SNE)
- Save results back to Excel and persist models with `joblib`

Files
- `gui.py` — Tkinter GUI and user workflows
- `cluster_tool.py` — clustering engine and utilities
- `run.bat` — Windows launcher that creates a `.venv` and installs `requirements.txt` on first run; displays an ASCII banner
- `ascii_banner.txt` — plain-text ASCII banner used by `run.bat`

Quickstart — Windows (recommended)
1. Double-click `run.bat` in the project root. On first run it will:
	- create a virtual environment `.venv`
	- install dependencies from `requirements.txt`
	- display the ASCII banner and launch the GUI

2. If you need to run manually from PowerShell or CMD:
```powershell
cd path\to\text-clustering-tool
.venv\Scripts\Activate.ps1   # PowerShell
or
.venv\Scripts\activate.bat   # CMD
.venv\Scripts\python.exe gui.py
```

Quickstart — macOS / Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
.venv/bin/python gui.py
```

Usage (GUI)
- Click **Select Excel file...** and pick a workbook
- Choose the sheet and text column
- Pick algorithm and parameters, then click **Run Clustering**
- Edit suggested cluster names and click **Save Results** to write an output Excel file

CLI (advanced)
See `cluster_tool.py` for a CLI entrypoint and example usage (column flags, algorithm and output options).

Support & Contact
- Website: https://aneekhait.github.io
- Author: Aneek Hait

License
This project is MIT licensed — see `LICENSE` for details.
