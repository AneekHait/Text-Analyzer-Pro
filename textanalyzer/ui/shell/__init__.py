"""Shell-level chrome for the Text Analyzer workspace.

Modules in this package (added incrementally):

- ``workspace_tabs``: The top-level :class:`WorkspaceTabWidget` that hosts
  one or more analysis sessions as tabs.
- ``dock_panels``: Optional left ("Navigator") and right ("Inspector") dock
  panels for recent files, open sessions, and selected-item details.

Phase 2 introduces this scaffolding without changing analysis behaviour:
the existing single-page UI is wrapped as the only initial tab.
"""

__all__ = ["workspace_tabs", "dock_panels"]
