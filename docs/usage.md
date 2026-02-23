# Usage and Notes

This document describes how to run the Text Clustering Tool (GUI and CLI).

GUI (`gui.py`)
- Launch the GUI and use the "Select Excel file..." button to choose an Excel (.xlsx) file.
- Select the text column from the dropdown, choose a clustering algorithm and parameters, then click "Run clustering".
- After clustering finishes, edit cluster names if desired and click "Save results" to write a new Excel file with `cluster_label` and `cluster_name` columns.

CLI (`cluster_tool.py`)
- The CLI supports the following arguments:
  - `--input` / `-i`: input Excel file (required)
  - `--column` / `-c`: text column name to cluster (required)
  - `--algorithm` / `-a`: `kmeans`, `dbscan`, or `agglomerative` (default: `kmeans`)
  - `--n_clusters` / `-k`: number of clusters for kmeans/agglomerative (default: 5)
  - `--output` / `-o`: output Excel path (defaults to `<input>_clustered.xlsx`)
  - `--visualize` / `-v`: produce a 2D visualization (PCA or t-SNE)

Examples

```bash
# Run kmeans with 5 clusters
"$PWD/.venv/bin/python" cluster_tool.py -i data.xlsx -c comments -a kmeans -k 5 -o data_clustered.xlsx

# Run DBSCAN (cosine metric)
"$PWD/.venv/bin/python" cluster_tool.py -i data.xlsx -c comments -a dbscan --eps 0.4 --min_samples 3
```

Dependencies

See `requirements.txt` for the Python libraries used. Install them into a virtualenv:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Notes
- Tkinter is required for the GUI. On Debian/Ubuntu: `sudo apt-get install python3-tk`.
- If your Excel file has multiple sheets, pass `--sheet` with a name or index.
- Visualization uses matplotlib and seaborn; large datasets may be slow.
