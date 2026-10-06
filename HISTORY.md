# History

Changes to Item Checker, newest first.

## 2026-10-06: Phone resolution adaptation & Item Search workflow (Desktop & APK)

- **Phone resolution responsiveness:**
  - Added `viewport-fit=cover` and safe-area inset padding for modern mobile phone screens.
  - Header navigation adapts responsively: on phone resolutions (360px–412px), the tabs become a full-width, touch-friendly segmented control right beneath the header bar for easy thumb switching between `Item Search` and `All Items`.
  - Item Search card scales dynamically across narrow screen widths without clipping or horizontal overflow.
  - **Auto-adjusting Search & Clear buttons:** In `Item Search`, the `[ Search ]` and `[ Clear ]` buttons automatically adapt and wrap on narrow/phone window widths, remaining centered and accessible without getting clipped.
  - Verifier count inputs support `inputmode="numeric"` to summon the phone's numeric keypad directly.
  - Stat cards in `All Items` cleanly form a 2x2 grid on mobile phones (Total Items, Active, To count, Mismatch).
- **Removed Physical Stock from Item Card:** Physical Stock is no longer visible on the item verification card/modal in both the desktop Tkinter app (`ItemChecker.pyw`) and mobile/APK app (`ItemChecker.html`), maintaining blind counting verification integrity.
- **Auto edit mode for empty counts:** When opening an item card, any verifier without a saved count (`None`/empty) opens directly in an entry field (`Edit mode`) with an inline Save button.
- **Locked counts with inline edit:** When a verifier has a saved count, it displays as a formatted number (highlighted in red if differing from physical stock) with an `Edit` button to unlock and modify it.
- **Fast keyboard workflow:** Pressing `<Enter>` in Verifier 1 saves the value and immediately moves focus to Verifier 2 if it is also in edit mode.
- **Instant result calculation:** Saving counts immediately recalculates and displays status (`MATCH`, `MISMATCH`, `TO COUNT`) in real time.
- **New "Item Search" tab & embedded card view:** Added a dedicated "Item Search" tab featuring a fast scan/search bar at the top. Searching by SKU (or scanning a barcode) immediately renders the item card directly inside the tab with direct count inputs and no Close button.
- **Renamed "Items" to "All Items":** The full imported item list, table view, stat summary counters, and filters are now located in the "All Items" tab.
- **Hidden SQL Console tab:** The SQL Console tab is hidden from the main tab interface.
- **Removed bottom Edit button:** Removed the bottom `Edit` button from the item card/modal. Each verifier has its own inline `Edit`/`Save` controls directly in its tile, keeping the bottom clean with only `Delete` (no Close button in embedded search card).
- **Android APK rebuilt:** `ItemChecker.apk` recompiled and synced with updated web assets.

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
