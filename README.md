# Bank of America Monthly Expense Tracker

A local Windows desktop application for turning Bank of America credit-card and debit-account PDF statements into a compact Excel expense report.

## Project requirements

The verified current-state product requirements are documented in [`requirements/`](requirements/README.md). The root `requirements.txt` remains the Python dependency file.

## Development and testing

Install the runtime and pytest dependencies, then run the isolated test suite:

```powershell
py -m pip install -r requirements-dev.txt
py -m pytest -q
```

Tests are stored in [`tests/`](tests/README.md) and trace covered behavior to the stable project requirement IDs. The initial suite covers `FR-001` without opening the real desktop interface or writing to the application's local data directories.

## Run it

1. Install Python 3.10 or newer.
2. In this folder, run `py -m pip install -r requirements.txt`.
3. Double-click `run_tracker.bat` (or run `py expense_tracker.py`).
4. Select one or more BofA statement PDFs and choose **Export automatically**.

The workbook is written to `output/expense_tracker.xlsx`.  The application stores its local import history and merchant rules in `data/tracker.sqlite3`; neither is uploaded anywhere.  Delete that database only if you deliberately want to start the local history over.

For repeatable testing, use **Reset ALL…** in the app. After confirmation it removes the local database and every generated file in `output/`, but never deletes your source PDFs. Use **Reset transaction history…** when you want to remove imported statements and generated Excel output while retaining your merchant/category rules and custom categories.

**Export automatically** exports familiar merchants immediately and places ordinary new merchants in `Other` under their merchant name. It requests input only for ambiguous activity (for example, cash withdrawals, person-to-person payments, transfers, Amazon/Google charges, or crypto) and proposes `Other → Merchant` as the safe fallback. Use **Review all new merchants** when you want to categorize every new merchant before export.

The app also recognizes common merchant patterns locally. For example, it categorizes Panera, Starbucks, DoorDash, and Osaka as **Restaurants**; Amazon, AliExpress, eBay, Etsy, and Walmart.com as **Online shopping**; and YouTube, Netflix, Spotify, and Hulu as **Entertainment & subscriptions**. Your own saved category choices always take priority.

## Correcting `Other` categories

After statements have been imported, choose **Edit Other categories…** on the main screen. Select a merchant from the list, choose an existing category or create a new one, enter its subcategory, and click **Queue assignment**. Nothing is changed until **Save changes and close** is selected. Then use **Export all saved months** to rebuild every Excel worksheet with the corrected totals; no PDF needs to be selected again.

Each monthly worksheet separates category/subcategory amounts by **Payment source**. Rows sourced from debit-account statements are shaded yellow; credit-card rows are unshaded.

## Notes

- Version 1 supports text-based BofA statements. Password-protected and scanned/image-only PDFs are reported without changing the workbook.
- A worksheet is named after the statement closing month (`YYYY-MM`). Importing the same month rebuilds that worksheet from the locally stored statement data.
- The workbook contains category totals only, not individual transaction rows.
