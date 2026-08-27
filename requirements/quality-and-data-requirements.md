# Quality and data requirements

## Quality requirements

- **QR-001 — Supported runtime:** The supported user environment shall be Windows with Python 3.10 or newer and Tkinter available in the Python installation.
- **QR-002 — Runtime dependencies:** Runtime PDF and workbook behavior shall use the dependency ranges declared in the root `requirements.txt`: `pypdf>=5.0,<7.0` and `openpyxl>=3.1,<4.0`.
- **QR-003 — Local-only operation:** Statement extraction, transaction processing, categorization, SQLite storage, and workbook generation shall run locally. The application shall not transmit financial data to external services.
- **QR-004 — Deterministic results:** Given the same parser version, statement bytes, saved rules, and local data state, parsing identities and category totals shall be reproducible.
- **QR-005 — Batch reliability:** A parsing failure in any selected statement shall prevent statement/transaction persistence and workbook generation for the entire selected batch.
- **QR-006 — User-visible errors:** Expected input failures shall be converted to understandable dialog messages instead of exposing an unhandled parser exception to the user.
- **QR-007 — Backward-compatible storage:** Startup shall preserve existing SQLite data while adding the supported `parser_version` and `source_type` columns when an older database lacks them.
- **QR-008 — Parser compatibility:** A change that alters imported rows, fingerprints, source types, or statement interpretation shall increment `PARSER_VERSION`, allowing a previously seen statement to be replaced using the new parser behavior.
- **QR-009 — Maintainability:** Monetary and parsing behavior should remain separable from Tkinter callbacks, use type annotations, and follow the repository conventions in `AGENTS.md`.
- **QR-010 — Test isolation:** Automated tests shall redirect `DATA_DIR`, `OUTPUT_DIR`, `DATABASE`, and `WORKBOOK` to temporary paths and shall use synthetic financial inputs. Tests shall never read or modify the user's real local data or output.
- **QR-011 — Baseline limits:** This current-state specification defines no response-time, throughput, concurrency, availability, remote-backup, or recovery-time target.

## Data requirements

- **DR-001 — Application paths:** Persistent application state shall be stored under `data/`, and generated workbook output shall be stored under `output/`, relative to the application file.
- **DR-002 — Database contents:** SQLite shall maintain statement identity and parser version, transaction identity and source data, saved merchant rules, and custom category names.
- **DR-003 — Account minimization:** The application shall retain only the final four digits extracted from a statement account number for application-level account identification.
- **DR-004 — Monetary precision:** Parsed and aggregated money shall use `Decimal`; SQLite shall store transaction amounts as decimal strings, and Excel shall receive numeric decimal values.
- **DR-005 — Source files:** The application shall read and hash selected source PDFs but shall not copy them into `data/` or `output/`. Reset operations shall never delete the source PDFs.
- **DR-006 — Transaction descriptions:** The complete cleaned statement description shall be retained locally because it participates in transaction identity, categorization review, and merchant normalization.
- **DR-007 — Source classification:** Each stored transaction shall retain its credit, debit, or compatibility fallback source type so workbook output can distinguish payment sources.
- **DR-008 — User rules:** Saved merchant rules and custom category names shall persist across statement replacement, repeated export, and transaction-history reset. They shall be removed by **Reset ALL**.
- **DR-009 — Generated artifacts:** `data/`, `output/`, `input/`, Python bytecode, real statements, and generated workbooks shall remain outside version control.
- **DR-010 — Retention control:** Local state shall remain until the user invokes a confirmed reset or deletes it outside the application; the application performs no automatic expiration or cloud backup.
