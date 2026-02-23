# Text Clustering Tool (Text Analyzer Pro)

Python desktop GUI for **text clustering from Excel files** using **TF-IDF**, **KMeans**, **DBSCAN**, and **Agglomerative Clustering**.

Use this project for customer feedback clustering, survey response analysis, support ticket grouping, and general NLP text analysis on tabular data.

## Why this project

- Works directly with Excel workbooks (`.xlsx`) and multiple sheets
- Non-technical friendly GUI built with Tkinter
- Auto keyword extraction and human-readable cluster naming
- 2D cluster visualization with PCA or t-SNE
- Optional CLI for scripting and repeatable workflows

## Features

- Multi-sheet Excel support; pick sheet + text column
- Text preprocessing and TF-IDF vectorization
- Clustering algorithms: `kmeans`, `dbscan`, `agglomerative`
- Cluster keyword extraction and suggested cluster names
- Visualizations: PCA and t-SNE
- Save clustered output back to Excel
- Persist models with `joblib`

## Quickstart

### Windows (recommended)

1. Double-click `run.bat` in the project root.
2. On first run it creates `.venv`, installs `requirements.txt`, and launches the GUI.

Manual launch:

```powershell
cd path\to\text-clustering-tool
.venv\Scripts\Activate.ps1
.venv\Scripts\python.exe gui.py
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python gui.py
```

## GUI usage

1. Click **Select Excel file...**
2. Choose the sheet and text column
3. Select algorithm + parameters
4. Click **Run Clustering**
5. Edit suggested names (optional), then click **Save Results**

## CLI usage

```bash
# KMeans
python cluster_tool.py -i data.xlsx -c comments -a kmeans -k 5 -o data_clustered.xlsx

# DBSCAN
python cluster_tool.py -i data.xlsx -c comments -a dbscan --eps 0.4 --min_samples 3
```

More examples: `docs/usage.md`

## Project files

- `gui.py`: Tkinter GUI workflow
- `cluster_tool.py`: clustering engine + CLI
- `run.bat`: one-click Windows launcher
- `docs/usage.md`: detailed usage and notes

## Contributing

Contributions are welcome. Start with:

- `CONTRIBUTING.md`
- `.github/ISSUE_TEMPLATE/feature_request.yml`
- `.github/ISSUE_TEMPLATE/bug_report.yml`
- `CHANGELOG.md`

## Security

To report vulnerabilities, see `SECURITY.md`.

## Support

- Website: https://aneekhait.github.io
- GitHub Sponsors: https://github.com/sponsors/AneekHait
- Buy Me a Coffee: https://www.buymeacoffee.com/aneekhait
- LinkedIn: https://www.linkedin.com/in/aneekhait/

If this tool saved you time, details on funded feature requests are in `FUNDING.md`.

## Growth checklist

For maintainers: `docs/github-growth-checklist.md`

## License

MIT license. See `LICENSE`.
