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
- `ItemChecker.html`: web version of the checker (`index.html` redirects to it)
- `supabase_setup.sql`: tables and access rules for the shared online list
- `schema.sql`: SQLite schema (table, `updated_at` trigger, `v_item_check` view)
- `Item_Checker_Project_Plan.pdf`: project plan

## Running

Requires Python 3.11 or newer. Excel import/export is optional and needs `openpyxl`:

```
pip install openpyxl
pythonw ItemChecker.pyw
```

The app creates `items.db` next to the script on first run. To use a different location, set the `ITEMCHECKER_DB` environment variable. The database is not tracked in git.

## Shared online list (Supabase)

When `SUPABASE_URL` and `SUPABASE_ANON_KEY` are filled in, at the top of `ItemChecker.pyw` and in the `<script>` of `ItemChecker.html`, everyone works on **one shared list**:

- **Web:** https://yhelie08.github.io/Item-List/ (GitHub Pages). Share this link.
- **Desktop:** `ItemChecker.pyw` uses the same online list. It keeps a local copy in `items_online_cache.db` and never overwrites `items.db`.
- Everyone signs in with an email and password. Changes by one person show up for the others within about 5 seconds.
- If the online list is empty, the desktop app offers to upload the items from `items.db`.

### One-time setup

1. Create a free project at https://supabase.com.
2. In **SQL Editor**, paste all of `supabase_setup.sql` and click **Run**.
3. In **Authentication → Sign In / Providers**, turn off **Allow new users to sign up** so only people you add can get in.
4. In **Authentication → Users → Add user → Create new user**, add each person with an email and password. Tick **Auto Confirm User**.
5. In **Project Settings → API**, copy the **Project URL** and the **anon public** key into both files.
6. In the GitHub repo, go to **Settings → Pages**, set the source to **Deploy from a branch → main → / (root)**, and save.

The anon key is meant to be public. The row-level security rules in `supabase_setup.sql` only let signed-in users read or change items.

Without these settings, both apps work on their own like before: the desktop uses `items.db`, and the web page saves in the browser.
