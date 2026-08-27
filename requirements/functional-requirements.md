# Functional requirements

## Launch and statement selection

- **FR-001 — Windows launch:** The application shall start with Python 3.10 or newer by running `py expense_tracker.py` or `run_tracker.bat` from the repository root.
- **FR-002 — Desktop interface:** The application shall present a Tkinter desktop interface for selecting statements, importing and exporting data, reviewing categories, resetting local state, and opening the output folder.
- **FR-003 — PDF selection:** The statement picker shall accept one or more files filtered to the PDF file type and display the selected file names before import.
- **FR-004 — Selection required:** An import request without selected statements shall show an informational message and shall not import or export data.

## Statement validation and parsing

- **FR-005 — Supported statements:** The application shall support text-based Bank of America credit-card and debit-account PDF statements whose extracted layout contains a recognizable account number, statement period, and transaction section.
- **FR-006 — Unsupported PDFs:** The application shall reject password-protected PDFs, PDFs with insufficient extractable text, unrecognized account or statement-period layouts, missing transaction sections, and layouts from which no transaction rows can be parsed.
- **FR-007 — Helpful failures:** A rejected statement shall produce a user-facing reason, including guidance for locked or scanned/image-only documents when applicable.
- **FR-008 — Batch validation:** The application shall parse every selected statement before saving statements or transactions. If any selected statement fails, it shall save none of the selected statements or transactions and shall not update the workbook.
- **FR-009 — Statement identity:** The application shall derive the account identity from the final four account digits, derive the reporting month from the statement closing period in `YYYY-MM` form, and classify the statement source as credit or debit.
- **FR-010 — Transaction rows:** The parser shall recognize transaction rows containing a posting date, description, and two-decimal amount, while retaining eligible continuation text as part of the complete description.
- **FR-011 — Expense filtering:** Credit imports shall retain positive purchase amounts, debit imports shall retain negative subtraction amounts, and stored/reportable expenses shall use the absolute monetary value. Known internal payments and transfers shall be excluded.
- **FR-012 — Transaction identity:** Each transaction shall receive a deterministic SHA-256 fingerprint derived from its account, month, posting date, complete normalized source row, and expense amount. The complete row shall remain part of the identity so distinct purchases with the same merchant, date, and amount are not automatically collapsed.
- **FR-013 — Statement digest:** Each source PDF shall receive a SHA-256 digest based on its file bytes for duplicate-import detection.

## Categorization and review

- **FR-014 — Rule precedence:** A user-saved merchant rule shall take precedence over built-in exact and pattern-based suggestions.
- **FR-015 — Built-in suggestions:** The application shall use its local built-in rules to suggest high-confidence category and subcategory values for recognized merchants. All categorization shall occur locally.
- **FR-016 — Unknown merchants:** During automatic export, an ordinary unknown merchant shall be reported under `Other` with the normalized merchant name as its subcategory without creating a permanent merchant rule.
- **FR-017 — Ambiguous activity:** During automatic export, new ambiguous activity such as cash withdrawals, person-to-person payments, transfers, marketplace/platform charges, or crypto activity shall be presented for user review before statement data is saved.
- **FR-018 — Review-all mode:** The user shall be able to request review of every new merchant that lacks an applicable rule, rather than only ambiguous activity.
- **FR-019 — Review defaults:** The review dialog shall show the normalized merchant and transaction examples, default to `Other` and the title-cased merchant name, and require a non-empty subcategory before saving the rule.
- **FR-020 — Review cancellation:** Closing merchant review before completion shall cancel statement and transaction persistence for the batch. Merchant rules explicitly saved on earlier review steps remain persisted because rules are saved as the user advances.
- **FR-021 — Base categories:** The review flow shall offer the built-in categories: Groceries, Beauty, Car parts, Home cleaning, House bills, Restaurants, Shopping, Online shopping, Entertainment & subscriptions, Insurance, Fuel, Pet care, Alcohol, Healthcare, Home improvement, Dealership, Having fun, and Other.
- **FR-022 — Edit uncategorized merchants:** The user shall be able to review saved merchants that still resolve to `Other`, see their transaction count and total, queue category/subcategory assignments, and create a custom category name.
- **FR-023 — Deferred category edits:** Queued edits from the uncategorized-merchant dialog shall not be persisted until the user selects **Save changes and close**. Canceling the dialog shall discard queued edits.

## Persistence and import replacement

- **FR-024 — Local persistence:** The application shall persist statements, transactions, merchant rules, and custom categories in the local SQLite database at `data/tracker.sqlite3`.
- **FR-025 — Duplicate import:** Re-importing an identical statement with the current parser version shall add no transactions.
- **FR-026 — Replacement import:** Importing a different or newly parsed statement for the same account and closing month shall replace that account-month's previous statement and transactions while preserving other accounts and months.
- **FR-027 — Rule retention:** Replacing or re-importing statement data shall not remove user-saved merchant rules or custom categories.

## Excel output

- **FR-028 — Workbook location:** A successful import shall create or update `output/expense_tracker.xlsx`.
- **FR-029 — Monthly worksheets:** The workbook shall contain a worksheet named `YYYY-MM` for each exported closing month. Re-exporting a month shall replace that worksheet rather than append a duplicate.
- **FR-030 — Summary contents:** Each monthly worksheet shall contain category totals and subcategory rows split by payment source, plus an overall total. It shall not expose individual transaction rows.
- **FR-031 — Payment-source presentation:** Credit and debit expenses shall have distinct source labels. Debit detail rows shall use yellow shading; credit detail rows shall remain unshaded.
- **FR-032 — Money formatting:** Detail, category, and overall totals shall be written as numeric workbook values formatted as US-dollar currency.
- **FR-033 — Rebuild saved months:** The user shall be able to rebuild every saved month from SQLite without reselecting source PDFs. If no saved months exist, the application shall report that there is nothing to export.
- **FR-034 — Open output folder:** The user shall be able to create and open the output directory from the application.

## Reset and shutdown

- **FR-035 — Reset transaction history:** After explicit confirmation, **Reset transaction history** shall delete imported statements and transactions and remove generated output while retaining merchant rules and custom categories.
- **FR-036 — Reset all data:** After explicit confirmation, **Reset ALL** shall delete the application database, SQLite sidecar files when present, and generated output, then initialize a fresh store. It shall not delete source PDFs.
- **FR-037 — Reset cancellation:** Declining either reset confirmation shall leave local database and output data unchanged.
- **FR-038 — Clean shutdown:** Closing the application shall close the SQLite connection before destroying the main window.
