# Item Checker

A small desktop app for checking stock levels. It keeps an item list in a local SQLite database, compares **Physical Stock** against **Committed Stock**, and flags the items that need attention.

## Status rules

Rules are checked in order:

| Status | Rule |
| --- | --- |
| OVER-COMMITTED | committed > physical |
| NO STOCK | physical = 0 |
| LOW | available (physical − committed) ≤ 5 |
| OK | everything else |

## Files

- `ItemChecker.pyw`: the desktop app (Python + Tkinter)
- `ItemChecker.html`: browser version of the checker
- `schema.sql`: SQLite schema (table, `updated_at` trigger, `v_item_check` view)
- `Item_Checker_Project_Plan.pdf`: project plan

## Running

Requires Python 3.11 or newer. Excel import/export is optional and needs `openpyxl`:

```
pip install openpyxl
pythonw ItemChecker.pyw
```

The app creates `items.db` next to the script on first run. To use a different location, set the `ITEMCHECKER_DB` environment variable. The database is not tracked in git.

## Desktop and browser share the same data

Click **Open in Browser** in the desktop app (or **File → Open in browser**). This opens the HTML version at `http://127.0.0.1:8765/`, and it reads and saves the same `items.db`:

- Changes made in the browser show up in the desktop app within about 2 seconds, and the reverse also works.
- The pill in the browser header shows **Synced · items.db**. If the desktop app is closed, it turns red and changes are **not** saved.
- If you open `ItemChecker.html` directly (double-click), it shows **Browser only**. That copy keeps its data in the browser and does not touch `items.db`. To move that data over, use **Export → SQLite backup (.db)** there, then **Import** the file in the desktop app.
