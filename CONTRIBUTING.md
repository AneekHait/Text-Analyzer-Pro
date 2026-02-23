# Contributing Guide

Thanks for contributing to Text Clustering Tool.

## Ways to contribute

- Report bugs with reproducible steps
- Propose enhancements with a clear use case
- Improve docs and examples
- Submit code changes with tests where possible

## Development setup

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Run the GUI:

```bash
python gui.py
```

Run CLI help:

```bash
python cluster_tool.py --help
```

## Pull request checklist

- Keep PRs focused on one change
- Include a short "what changed and why"
- Update docs when behavior changes
- Avoid breaking existing CLI flags and GUI flows
- Keep naming and style consistent with project files

## Commit message style

Use clear, imperative messages, for example:

- `fix: handle empty text rows before vectorization`
- `feat: add silhouette score summary to output`
- `docs: add dbscan parameter examples`

## Issue labels

Suggested labels for maintainers:

- `bug`
- `enhancement`
- `good first issue`
- `documentation`
- `help wanted`
