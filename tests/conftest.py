"""Shared fixtures for isolated application tests."""

from pathlib import Path

import pytest

import expense_tracker


@pytest.fixture
def isolated_app_paths(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> dict[str, Path]:
    """Redirect generated application state to a per-test temporary directory."""
    data_dir = tmp_path / "data"
    output_dir = tmp_path / "output"
    paths = {
        "data": data_dir,
        "output": output_dir,
        "database": data_dir / "tracker.sqlite3",
        "workbook": output_dir / "expense_tracker.xlsx",
    }

    monkeypatch.setattr(expense_tracker, "DATA_DIR", paths["data"])
    monkeypatch.setattr(expense_tracker, "OUTPUT_DIR", paths["output"])
    monkeypatch.setattr(expense_tracker, "DATABASE", paths["database"])
    monkeypatch.setattr(expense_tracker, "WORKBOOK", paths["workbook"])

    return paths
