# History

Changes to Item Checker, newest first.

## 2026-09-30: Stock count with two verifiers

- **Committed Stock and Available are removed** from the form, item card, stat cards, import and export. Older databases keep the column, unused.
- **New field: Bin Location** (optional), with suggestions from bins already used. Search matches it. A Zoho export's `CF.Loc` column is imported as Bin Location (`Bin Name` is used only when there is no CF.Loc).
- **Add Item:** Item Name and SKU are still required.
- **Edit Item:** Item Name and SKU are read-only.
  - Physical Stock 0: Physical Stock can be edited.
  - Physical Stock more than 0: Physical Stock is read-only, and **Verifier 1 (Count)** and **Verifier 2 (Count)** appear.
- New status rules: **MATCH** (both counts equal Physical Stock), **MISMATCH**, **TO COUNT** (a count is missing), **NO STOCK**. The item card shows the count result; a count that differs is red.
- Stat cards: **To count** and **Mismatch** replace Committed QTY and Over-committed; tap them to filter the list.
- Export adds Bin Location, Verifier 1 Count, Verifier 2 Count and a read-only Count Result column. Import reads them; a file without those columns keeps the saved values.

## 2026-09-30: Android app

- New `android-app/` folder: Item Checker as an Android app, built from the same screens as `ItemChecker.html`.
- Works without internet; its libraries are saved inside the app. The list is saved on the phone.
- Import shows every file so `items.db` can be picked; Export opens the phone's Share menu.
- Install file: `ItemChecker.apk`. How to install and rebuild is in the README.

## 2026-09-30: Warehouse field and Zoho Inventory import

- New optional **Warehouse** field in both apps, Supabase and the SQLite schema. It shows on the item card and in the Add/Edit form (with suggestions from warehouses already used), and search matches it. The list still shows only Item Name, SKU and ItemRef.
- Included in Excel/CSV import and export and the import template. Importing a file without a Warehouse column keeps the warehouses already saved.
- Import reads a Zoho Inventory item export directly: `Warehouse Name` → Warehouse, `CF.ITEMREF` → ItemRef, `Stock On Hand` → Physical Stock, `Status` (Active/Inactive) → Active.
- Supabase needs the column `warehouse` (exact name). `supabase_setup.sql` adds it.

**Shared online list turned off**
- Supabase is no longer used. The desktop app works on `items.db` again, and the web page saves in the browser.
- The shared-list code is still there and comes back when the Supabase URL and key are filled in again.

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
