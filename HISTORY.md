# History

Changes to Item Checker, newest first.

## 2026-09-29: ItemRef field and a simpler, responsive item view

**New field: ItemRef**
- Optional text field on every item, in the web page, the desktop app, Supabase and the SQLite schema.
- Included in search, sorting, the Add/Edit form, Excel/CSV import and export, and the import template.
- Importing a file without an ItemRef column keeps the ItemRefs already saved.
- Older databases (`items.db`, browser storage, `.db` backups) get the column automatically when opened.
- Supabase needs the column `item_ref` (exact name). `supabase_setup.sql` adds it.

**Item list shows only Item Name, SKU and ItemRef**
- Opening an item shows a card with Item Name, SKU, ItemRef, Physical Stock, Committed Stock, Available and Active, plus Edit, Mark active/inactive and Delete.
- Web: tap or click the item. Desktop: double-click, or select it and press Enter or View.
- The desktop table no longer colors rows by status. Inactive items are still greyed out, and the status filter still works.

**Works on phones and small windows**
- Web: a phone gets a tappable list, a round + button to add items, a card that slides up from the bottom, and inputs that do not zoom the page.
- Desktop: the columns stretch with the window; the stat cards, header buttons, filters and action buttons wrap on a narrow window (down to 380 px wide).

## 2026-09-25: Shared online list

- Both apps share one item list through Supabase, with email/password sign-in.
- The web version is published on GitHub Pages: https://yhelie08.github.io/Item-List/
- The desktop app keeps its online copy in `items_online_cache.db` and never overwrites `items.db`.

## 2026-09-25: First version

- Desktop app (Python + Tkinter) and web version, both on SQLite.
- Status rules: OVER-COMMITTED, NO STOCK, LOW, OK.
- Import and export as Excel, CSV and SQLite backup; SQL console.
