# Changelog

All notable changes to this project are documented here. The format is based
on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and this project
adheres to [Semantic Versioning](https://semver.org/).

## [2.0.0] — 2026-05

A full reimagining of Text Analyzer Pro. The Tk-era 1.x line is replaced by a
modular **PySide6** desktop app with a port of every feature that previously
only existed in the older `Text-Analyzer-Pro 1.5` codebase, plus a
suite of new productivity tools and a rewritten Wordcloud Studio.

### Architecture & internals

- **Full PySide6 rewrite.** The legacy Tk GUI is gone. New entry point still
  `gui.py`, but the implementation is now modular under `textanalyzer/`:
  controllers, services, models, workers, ui, components, shell, utils.
- **`textanalyzer.engine` package.** `cluster_tool.py` and `wordcloud_tool.py`
  moved into `textanalyzer/engine/cluster.py` and `wordcloud.py`. The original
  root files are now ~10-line `sys.modules` swap shims so external scripts
  using `from cluster_tool import ...` keep working unchanged.
- **`textanalyzer.settings` package.** `app_settings.py` moved alongside;
  same shim pattern preserves backward compatibility.
- **Theme system.** Token-driven QSS (`theme/tokens.py`, `theme/qss.py`) with
  light + dark modes, runtime asset generation for combobox chevrons that
  matches the active palette.
- **Worker threads** for clustering, embedding, and wordcloud generation —
  the UI never blocks on CPU work.
- **Test suite.** 116 tests under `tests/` (plus 10 legacy unittest tests at
  the repo root) running headlessly via `QT_QPA_PLATFORM=offscreen`.

### File-format support

- **All Excel variants now load** out of the box: `.xlsx`, `.xlsm`, `.xltx`,
  `.xltm`, `.xls`, `.xlsb`. Plus `.ods` (OpenDocument), `.csv`, `.json`.
- **Encoding fallback** for CSV/JSON: tries UTF-8 → UTF-8-with-BOM → cp1252 →
  latin-1, returning the first that decodes cleanly. No more
  `UnicodeDecodeError: 'utf-8' codec can't decode byte 0x96 …` on
  Excel-exported CSVs. Emits a one-line warning when the source isn't UTF-8.
- New `EXCEL_INPUT_EXTENSIONS` constant; pandas auto-picks the right engine
  per format (openpyxl / xlrd / pyxlsb / odfpy).

### Setup tab

- **Drag-and-drop dropzone** as the empty state. Supports drag from Explorer
  or click-to-browse with a multi-format file dialog filter.
- **Live data preview** — 5-row table appears once a file loads, with the
  selected text column's header bolded.
- **Algorithm dropdown reordered** for safety: `kmeans++` → `hdbscan` →
  `agglomerative` → `dbscan`. The DBSCAN entry has a tooltip suggesting
  HDBSCAN as the better default for TF-IDF data.
- **Suggest K button** — runs silhouette analysis over k=2..15 and recommends
  a cluster count with a confidence band (`high` / `medium` / `low`).
- **Elbow button** — same UI, but uses inertia-bend detection instead of
  silhouette. Better for noisy text data where silhouette is uniformly low.
- **Compare Algorithms button** — runs KMeans + DBSCAN + Agglomerative on the
  same matrix in a worker thread, shows a metrics table (silhouette,
  Calinski-Harabasz, Davies-Bouldin, runtime), and offers a "Use {best}"
  one-click action.
- **Advanced TF-IDF dialog** — exposes `max_features`, `min_df`, `max_df`,
  `ngram_range` (uni / bi / tri), and `HashingVectorizer` for memory-bounded
  runs on very large corpora.

### Cleaning tab

- **Lemmatize words (NLTK)** toggle. Reduces words to their base form
  (`running` → `run`); auto-downloads WordNet + Punkt corpora on first use.
- **Custom regex find/replace** for project-specific cleanups.
- **Cleaning recipes** — save / load / delete named cleaning configurations
  across sessions.
- **Live preview** of cleaning effects before clustering.

### Wordcloud Studio

- **Live preview** with debounced regen (350 ms). Change a setting; the cloud
  refreshes automatically.
- **Column selector** in the sidebar — switch text columns without leaving
  the dialog.
- **Contrast-aware gradient sampler** — colormaps are clipped to a
  background-aware sub-range so low-frequency words never disappear into the
  background. White bg gets the dark end of `Blues`; black bg gets the light
  end; medium bg compresses both.
- **Robust Copy to Clipboard** — round-trips through PNG bytes via
  `QImage.fromData`, paste-tested in PowerPoint, Word, and image editors.
  RGBA images are flattened onto white so apps that ignore alpha don't paste
  a black background.
- **All-formats save** — PNG, JPG, SVG via the toolbar.
- **Sidebar polish** — `WrapLongRows` form layout, shrinkable combos, fits at
  any sidebar width without clipping fields or labels.
- Branded "Accenture Purple" palette removed; default selection is now
  `Corporate Blue`.
- Minimize / maximize buttons added to the dialog frame.

### Clustering engine

- **HDBSCAN support** — auto-detects cluster count, no `n_clusters` needed.
  Optional dependency; clear ImportError with install hint when missing.
- **MiniBatchKMeans auto-switch** at 10k+ rows for memory-bounded runs on
  large corpora. Threshold tunable via `--minibatch_threshold`.
- **kmeans++ initialization** is now passed explicitly to all
  KMeans/MiniBatchKMeans construction sites (was already the sklearn default;
  now explicit and visible at the call site). Surfaced in the GUI dropdown
  as `kmeans++` while the engine still accepts the bare `kmeans` string.
- **Optional GPU acceleration** via cuML (`--use_gpu`) with silent CPU
  fallback when cuML isn't installed or fails to initialize.
- **`find_optimal_k(X, k_range, method)`** — silhouette- or elbow-based
  cluster-count search, with early-stop when scores decline for 3+
  consecutive k values.
- **`compare_algorithms(X, ...)`** — runs all CPU algorithms, returns a
  list of `AlgorithmResult` dataclasses with full metrics + runtime; skips
  agglomerative on >10k samples to avoid OOM.
- **`validate_input(df, column, ...)`** — pre-flight stats: total rows, null
  count, empty %, unique-text count, average / min / max text length, and
  algorithm-specific warnings (DBSCAN noise hint, agglomerative memory
  warning, n_clusters > unique_texts warning).
- **`ApplicableModel` wrapper** — gives DBSCAN / Agglomerative / HDBSCAN a
  working `.predict()` via a `NearestNeighbors` index built over the training
  matrix. Lets non-KMeans models be saved and re-applied to new data.
- **HashingVectorizer + chunked TF-IDF** for memory-bounded runs on huge
  corpora. `get_top_keywords_per_cluster` degrades gracefully when given a
  hashing vectorizer (returns empty keyword lists, warns once).
- **Fine-grained TF-IDF knobs**: `min_df`, `max_df`, `ngram_range`,
  `custom_stopwords` exposed on `vectorize_texts`.
- **`TextCleaningConfig.lemmatize` and `custom_stopwords`** fields, plumbed
  through `clean_text_value`.

### File menu

- **Save Model** bundles model + vectorizer + cluster names + top-keyword
  summary into a single `.joblib`. Non-KMeans models are auto-wrapped with
  `ApplicableModel` so the saved artifact has a working `.predict()`.
- **Load Model** opens a saved `.joblib`, transforms the current dataframe
  through the saved vectorizer, predicts via the saved model, jumps to the
  Results tab. Adopts the loaded artifacts as current run state so Save
  Results / Visualize / Save Model continue to work as if you'd just
  trained.

### CLI

- New flags: `--load_model`, `--use_hashing`, `--chunk_size`, `--use_gpu`,
  `--minibatch_threshold`, `--min_cluster_size`. `hdbscan` added to the
  `--algorithm` choices.
- `--load_model --model_path X --input Y` applies a saved model to fresh
  data and writes a `_clustered_from_model.{xlsx|csv|json}` output.
- Save-model now wraps non-KMeans models with `ApplicableModel` so the saved
  joblib remains useful for headless application later.

### Performance

- **Lazy imports** of matplotlib, seaborn, and NLTK in the engine module.
  Cold start of `import textanalyzer.engine.cluster` dropped by **~1 second**
  on Windows.
- **`_to_dense(X)` helper** consolidates 8 occurrences of the
  `X.toarray() if hasattr(X, "toarray") else X` idiom into one place.

### UX polish

- Both the main window and Wordcloud Studio open **maximized**.
- About dialog hugs its content; primary button takes initial focus.
- After loading a file, the user **stays on Setup** instead of being
  bounced to Cleaning automatically.
- The single-tab `WorkspaceTabWidget` auto-hides its tab bar; reappears when
  Phase-3 multi-document support adds a second tab.
- Every menu item, log message, and dialog title now uses Unicode glyphs
  correctly (mojibake cleanup across `gui.py`).
- Settings persist a "last algorithm" via the engine value, with backward
  compatibility for the old `kmeans` literal.

### Repository hygiene

- `.gitignore` expanded to cover `_clustered.*` outputs, `.joblib` models,
  `wordcloud_*.png` exports, `_vis_pca.png` / `_vis_tsne.png` snapshots,
  pytest / mypy / ruff caches, more venv variants, IDE folders, and OS junk.
- `requirements.txt` pins `scipy`, `xlrd`, `pyxlsb`, `odfpy`, `hdbscan`.
- All commits attributed to a single canonical author identity.

### Tests

- **116 pytest tests** under `tests/` covering: data-source panel state
  machine, main-window construction, dropzone path filtering, wordcloud
  dialog (column selector + live render plumbing + setup buttons), cluster
  tool ports (`validate_input`, `find_optimal_k`, `compare_algorithms`,
  `ApplicableModel`, hashing vectorizer, MiniBatchKMeans switching, HDBSCAN
  routing, encoding fallback, all Excel formats), and the GUI Suggest K /
  Elbow / Compare buttons.
- All tests run headlessly with `QT_QPA_PLATFORM=offscreen`.

### Dependencies

- **Added**: `scipy`, `xlrd`, `pyxlsb`, `odfpy`, `hdbscan`.
- **Already present**: `pandas`, `scikit-learn`, `openpyxl`, `matplotlib`,
  `seaborn`, `joblib`, `Pillow`, `PySide6`, `qtawesome`, `wordcloud`.

### Migration notes (1.x → 2.0)

- **No code changes required** for users importing from `cluster_tool`,
  `wordcloud_tool`, or `app_settings` — those modules still resolve to the
  same canonical objects.
- New code should prefer `from textanalyzer.engine.cluster import ...`,
  `from textanalyzer.engine.wordcloud import ...`,
  `from textanalyzer import settings`.
- Saved `.joblib` model files from v1.x continue to load via the **Load
  Model** action / `--load_model` CLI flag. KMeans models predict natively;
  non-KMeans models from v1.x without an `ApplicableModel` wrapper will
  raise a clear "model has no `.predict()`" error — re-save with v2.0 to
  add the wrapper.
- Cleaning recipes saved by v1.x load fine; the new `lemmatize` toggle
  defaults to off when missing from the persisted dict.

---

## [1.5] — earlier

- Wordcloud generation (initial release)
- TTK widgets for GUI scaffolding
- Excel multi-sheet support, CSV / JSON loading
- KMeans / DBSCAN / Agglomerative clustering
- TF-IDF vectorization with English stopwords
- 2D visualization (PCA / t-SNE)
- Save trained model (joblib)

## [Unreleased]

- (intentionally empty — pending the next round of changes)
