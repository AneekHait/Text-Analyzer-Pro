# Text Clustering Tool

A small GUI and CLI tool to cluster text data from Excel files using TF-IDF and scikit-learn.

Features
- Load an Excel file and choose a text column (GUI).
- Preprocess and vectorize texts (TF-IDF).
- Clustering algorithms: KMeans, DBSCAN, Agglomerative.
- Compute top keywords per cluster and generate human-readable cluster names.
- Save results back to Excel with cluster labels and names.

Files
- `gui.py` — a Tkinter GUI wrapper around the clustering utilities.
- `cluster_tool.py` — clustering utilities and a CLI entrypoint.
- `cluster.txt` — original clustering module source (copied into `cluster_tool.py`).

Quickstart (GUI)
1. Ensure system Tk support is installed (e.g., `sudo apt-get install python3-tk` on Debian/Ubuntu).
2. Install Python dependencies into your environment (recommended: virtualenv):

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

3. Launch the GUI:

```bash
"$PWD/.venv/bin/python" "${PWD}/gui.py"
```

Quickstart (CLI)

```bash
# Example: cluster the column "comments" in data.xlsx
"$PWD/.venv/bin/python" cluster_tool.py -i data.xlsx -c comments -a kmeans -k 5 -o data_clustered.xlsx
```

Documentation
See `docs/usage.md` for more examples and explanation of parameters.

License
This project is provided under the MIT license — see `LICENSE`.
