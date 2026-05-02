# Usage and Notes

This document describes how to run the Text Clustering Tool (GUI and CLI).

## GUI (`gui.py`)

### Getting Started
- Launch the GUI and use the **Select File…** button (or `Ctrl+O`) to choose an Excel, CSV, or JSON file.
- Select the text column from the dropdown, choose a clustering algorithm and parameters, then click **Run clustering** (`Ctrl+R`).
- After clustering finishes, edit cluster names if desired and click **Save results** (`Ctrl+S`) to write a new Excel, CSV, or JSON file with `cluster_label` and `cluster_name` columns.
- Use **Generate Wordcloud** after loading any sheet to open the dedicated wordcloud builder for the active column.

### Workspace Tabs
- Multiple analysis sessions can run side-by-side in tabs.
- Right-click a tab for context actions: **Close**, **Close Others**, **Close to Right**, **Duplicate**, **Rename**.
- Navigate tabs with `Ctrl+Tab` / `Ctrl+Shift+Tab` or close the active tab with `Ctrl+W`.

### Dock Panels
- **Navigator** (left) — quick access to recent files.
- **Inspector** (right) — context-sensitive details about the active item.
- Toggle docks from the **View** menu; layout persists across restarts.

### Secondary Windows
- **Settings** (`Ctrl+,`) — theme, defaults, cleaning data, shortcuts overview.
- **Diagnostics** (`Ctrl+/`) — filterable event log with Copy and Clear.
- **About** (`F1`) — version, credits, links.

### Keyboard Shortcuts
| Action               | Windows / Linux      | macOS              |
|----------------------|----------------------|--------------------|
| Open file            | `Ctrl+O`             | `⌘+O`             |
| Run clustering       | `Ctrl+R`             | `⌘+R`             |
| Save results         | `Ctrl+S`             | `⌘+S`             |
| Save model           | `Ctrl+Shift+S`       | `⌘+⇧+S`           |
| Close tab            | `Ctrl+W`             | `⌘+W`             |
| Next tab             | `Ctrl+Tab`           | `⌃+Tab`           |
| Previous tab         | `Ctrl+Shift+Tab`     | `⌃+⇧+Tab`         |
| Settings             | `Ctrl+,`             | `⌘+,`             |
| Diagnostics          | `Ctrl+/`             | `⌘+/`             |
| About / Help         | `F1`                 | `F1`               |
| Cancel running task  | `Esc`                | `Esc`              |

### Theme & Appearance
- Toggle between light and dark themes via the menu or the Settings window.
- Theme preference persists across sessions (stored in JSON and mirrored to OS settings).
- Font stack is platform-aware: Segoe UI Variable on Windows, SF Pro Text on macOS, Inter/Ubuntu on Linux.

### HiDPI & Multi-Monitor
- HiDPI scaling is enabled automatically (Qt `PassThrough` rounding policy).
- Window geometry is saved and restored on restart; if a saved monitor is disconnected the window is clamped to the nearest visible screen.

### Wordcloud Builder
- In the wordcloud builder, adjust max words, min frequency, output size, phrase mode, stopwords, normalization, and styling before clicking **Generate Preview**.
- Save the rendered cloud as PNG or export the filtered term-frequency table as Excel.

### Persistence
- Settings are stored in `~/.text_analyzer_pro/settings.json`.
- Selected keys (theme, geometry, dock state) are also mirrored to the OS native settings store (Windows Registry / macOS plist / XDG) for integration resilience.

---

## CLI (`cluster_tool.py`)
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

---

## Cross-Platform Notes

### Windows
- Uses Segoe UI Variable (or Segoe UI) as default font.
- Settings mirrored to `HKCU\Software\Aneek Hait\Text Analyzer Pro` in the Windows Registry.
- Native file dialogs used for Open / Save.

### macOS
- Uses SF Pro Text as default font.
- Application display name set for the About menu integration.
- Settings mirrored to `~/Library/Preferences/com.aneekhait.dev.Text Analyzer Pro.plist`.

### Linux
- Uses Inter, Ubuntu, or Cantarell — whichever is installed first.
- Desktop filename set to `text-analyzer-pro` for Wayland/X11 taskbar grouping.
- If no preferred font is found, falls back to the system default sans-serif.

---

## Acceptance Checklist

Use this checklist to verify full functionality after a fresh install:

- [ ] Toggle theme — instant, persists across restart.
- [ ] Close + reopen app — geometry, dock layout, last active tab restored.
- [ ] `Esc` cancels a running cluster job; Cancel button hides afterwards.
- [ ] All primary actions reachable by keyboard shortcuts.
- [ ] No `QMessageBox` flashes on success paths (only toasts).
- [ ] Resize window from 1280×880 down to 1040×740 — no layout breakage; scrollbars appear cleanly.
- [ ] WordCloud render is cancellable; empty filter state shown, not a stale image.
- [ ] Settings window edits persist; affect open tabs live.
