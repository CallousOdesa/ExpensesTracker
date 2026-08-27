"""Requirement tests for FR-002: the real Tkinter desktop interface."""

from __future__ import annotations

import sys
import tkinter as tk
from collections.abc import Callable, Iterator
from pathlib import Path
from tkinter import ttk

import pytest

import expense_tracker


BUTTON_CALLBACKS = {
    "Select PDF statements": "choose_files",
    "Export automatically": "export_automatically",
    "Export all saved months": "export_all_months",
    "Reset ALL": "reset_all",
    "Reset transaction history": "reset_transaction_history",
    "Edit Other categories": "edit_other_categories",
    "Review all new merchants": "review_and_export",
    "Open output folder": "open_output",
}

pytestmark = [
    pytest.mark.requirement("FR-002"),
    pytest.mark.gui,
    pytest.mark.windows,
    pytest.mark.skipif(sys.platform != "win32", reason="FR-002 requires Windows Tkinter"),
]


def descendants(widget: tk.Misc) -> Iterator[tk.Misc]:
    """Yield every child widget below a Tkinter widget."""
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


def stable_button_label(button: ttk.Button) -> str:
    """Return button wording without its optional typographic ellipsis."""
    return str(button.cget("text")).rstrip("\N{HORIZONTAL ELLIPSIS}\N{REPLACEMENT CHARACTER}")


@pytest.fixture
def app_factory(
    isolated_app_paths: dict[str, Path],
) -> Iterator[Callable[[], expense_tracker.App]]:
    """Construct hidden real applications and reliably release their resources."""
    applications: list[expense_tracker.App] = []

    def create() -> expense_tracker.App:
        app = expense_tracker.App()
        app.withdraw()
        applications.append(app)
        return app

    yield create

    for app in reversed(applications):
        app.store.close()
        app.destroy()


def test_fr_002_window_presents_identity_status_and_transaction_list(
    app_factory: Callable[[], expense_tracker.App],
) -> None:
    """The actual desktop window exposes its identity and initial state."""
    app = app_factory()

    widgets = list(descendants(app))

    assert app.title() == "BofA Monthly Expense Tracker"
    assert app.status.get() == "Select one or more Bank of America statement PDFs."
    assert app.listbox in widgets
    assert isinstance(app.listbox, tk.Listbox)
    assert app.listbox.winfo_manager() == "pack"
    assert any(
        isinstance(widget, ttk.Label)
        and str(widget.cget("textvariable")) == str(app.status)
        for widget in widgets
    )


def test_fr_002_window_presents_all_required_workflow_controls(
    app_factory: Callable[[], expense_tracker.App],
) -> None:
    """The real window contains every control required by FR-002."""
    app = app_factory()

    button_labels = {
        stable_button_label(widget)
        for widget in descendants(app)
        if isinstance(widget, ttk.Button)
    }

    assert button_labels == set(BUTTON_CALLBACKS)


def test_fr_002_workflow_controls_invoke_their_intended_callbacks(
    monkeypatch: pytest.MonkeyPatch,
    app_factory: Callable[[], expense_tracker.App],
) -> None:
    """Each real Tkinter button dispatches to the intended application callback."""
    invoked: list[str] = []

    def callback_spy(name: str) -> Callable[[expense_tracker.App], None]:
        def callback(app: expense_tracker.App) -> None:
            invoked.append(name)

        return callback

    for callback_name in BUTTON_CALLBACKS.values():
        monkeypatch.setattr(
            expense_tracker.App,
            callback_name,
            callback_spy(callback_name),
        )

    app = app_factory()
    buttons = {
        stable_button_label(widget): widget
        for widget in descendants(app)
        if isinstance(widget, ttk.Button)
    }

    for label in BUTTON_CALLBACKS:
        buttons[label].invoke()

    assert invoked == list(BUTTON_CALLBACKS.values())
