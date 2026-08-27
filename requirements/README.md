# Project requirements

This folder describes the verified, current behavior of the Bank of America Monthly Expense Tracker. It is intended for maintainers, reviewers, and test authors who need a stable product contract without reverse-engineering the application.

## Documents

- [Functional requirements](functional-requirements.md) describes user workflows and observable application behavior.
- [Quality and data requirements](quality-and-data-requirements.md) describes platform, privacy, reliability, compatibility, and data-handling constraints.

The root [`requirements.txt`](../requirements.txt) is the Python dependency file. It is separate from these product requirements.

## Scope and sources of truth

These documents are an as-built baseline, not a roadmap. They describe behavior verified in [`expense_tracker.py`](../expense_tracker.py) and the user guidance in the root [`README.md`](../README.md). They do not promise unsupported statement formats, services, performance targets, or future features.

When the documents and implementation disagree, the implementation identifies the current runtime behavior. Resolve the mismatch in the same change by either correcting the implementation or updating the applicable requirements and user documentation.

## Requirement identifiers

- `FR-*`: functional behavior visible through application workflows or generated output.
- `QR-*`: platform, privacy, reliability, compatibility, and maintainability qualities.
- `DR-*`: storage, retention, and financial-data handling.

Keep identifiers stable so they can be referenced by tests, issues, and reviews. Add new identifiers instead of renumbering existing requirements.

## Test coverage

Requirement tests live in [`tests/`](../tests/README.md). Test modules and functions include the applicable requirement ID and use the registered `requirement(id)` pytest marker. Current automated coverage includes [`FR-001`](../tests/test_fr_001_windows_launch.py) launch behavior and the real Tkinter interface required by [`FR-002`](../tests/test_fr_002_desktop_interface.py).
