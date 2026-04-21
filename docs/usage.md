# Usage and Notes

This document describes how to run the Text Clustering Tool (GUI and CLI).

GUI (`gui.py`)
- Launch the GUI and use the "Select File..." button to choose an Excel, CSV, or JSON file.
- Select the text column from the dropdown, choose a clustering algorithm and parameters, then click "Run clustering".
- After clustering finishes, edit cluster names if desired and click "Save results" to write a new Excel, CSV, or JSON file with `cluster_label` and `cluster_name` columns.
- Use "Generate Wordcloud" after loading any sheet to open the dedicated wordcloud builder for the active column.
- In the wordcloud builder, adjust max words, min frequency, output size, phrase mode, stopwords, normalization, and styling before clicking "Generate Preview".
- Save the rendered cloud as PNG or export the filtered term-frequency table as Excel.

CLI (`cluster_tool.py`)
- The CLI supports the following arguments:
  - `--input` / `-i`: input table file (`.xlsx`, `.xls`, `.csv`, `.json`) (required)
  - `--column` / `-c`: text column name to cluster (required)
  - `--algorithm` / `-a`: `kmeans`, `dbscan`, or `agglomerative` (default: `kmeans`)
  - `--n_clusters` / `-k`: number of clusters for kmeans/agglomerative (default: 5)
  - `--output` / `-o`: output path (defaults to `<input>_clustered` with a matching supported extension)
  - `--visualize` / `-v`: produce a 2D visualization (PCA or t-SNE)

Examples

```bash
# Run kmeans with 5 clusters
"$PWD/.venv/bin/python" cluster_tool.py -i data.xlsx -c comments -a kmeans -k 5 -o data_clustered.xlsx

# Run DBSCAN (cosine metric)
"$PWD/.venv/bin/python" cluster_tool.py -i data.xlsx -c comments -a dbscan --eps 0.4 --min_samples 3

# Run kmeans on CSV input
"$PWD/.venv/bin/python" cluster_tool.py -i data.csv -c comments -a kmeans -k 5 -o data_clustered.csv
```

Dependencies

See `requirements.txt` for the Python libraries used. Install them into a virtualenv:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Notes
- PySide6 is required for the GUI and is installed via `requirements.txt`.
- If your Excel file has multiple sheets, pass `--sheet` with a name or index. CSV and JSON inputs are treated as a single table.
- Visualization uses matplotlib and seaborn; large datasets may be slow.
- Wordcloud previews require the `wordcloud` package from `requirements.txt`.
