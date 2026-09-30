# Item Checker

A small desktop app for checking stock levels. It keeps an item list in a local SQLite database, and checks the **Physical Stock** of each item against the counts of two verifiers.

## Status rules

Rules are checked in order:

| Status | Rule |
| --- | --- |
| NO STOCK | physical = 0 |
| TO COUNT | Verifier 1 or Verifier 2 has not counted yet |
| MATCH | both verifier counts equal Physical Stock |
| MISMATCH | everything else |

## Item fields

Item Name, SKU, ItemRef (optional), Warehouse (optional), Bin Location (optional), Physical Stock QTY, Verifier 1 (Count), Verifier 2 (Count), Active. SKUs are unique and not case-sensitive.

**Add Item:** Item Name and SKU are required; the verifier counts are not shown.

**Edit Item:**
- Item Name and SKU are read-only.
- Physical Stock **0**: type the count into Physical Stock.
- Physical Stock **more than 0**: Physical Stock is read-only; enter Verifier 1 (Count) and Verifier 2 (Count). A blank count means not counted yet.
- To correct a Physical Stock that is already more than 0, import a file or use the SQL Console.

## Using the list

- The list shows only **Item Name, SKU and ItemRef**. Search matches those, Warehouse and Bin Location.
- Open an item to see its card: Item Name, SKU, ItemRef, Warehouse, Bin Location, Physical Stock, Verifier 1, Verifier 2, Count result and Active, with Edit, Mark active/inactive and Delete.
  - **Desktop:** double-click the item, or select it and press Enter or **View**.
  - **Web / phone:** tap the item. On a phone, add items with the round **+** button.
- Both apps adjust to the screen size: the web page fits phones, and the desktop window can be made as narrow as 380 px.

## Files

- `ItemChecker.pyw`: the desktop app (Python + Tkinter)
- `ItemChecker.html`: web version of the checker (`index.html` redirects to it). The Android app uses the same screens.
- `android-app/`: the Android app (Capacitor)
- `supabase_setup.sql`: tables and access rules for the shared online list (turned off for now)
- `schema.sql`: SQLite schema (table, `updated_at` trigger, `v_item_check` view)
- `Item_Checker_Project_Plan.pdf`: project plan
- `HISTORY.md`: list of changes

## Running

Requires Python 3.11 or newer. Excel import/export is optional and needs `openpyxl`:

```
pip install openpyxl
pythonw ItemChecker.pyw
```

The app creates `items.db` next to the script on first run. To use a different location, set the `ITEMCHECKER_DB` environment variable. The database is not tracked in git.

## Android app

`android-app/` packs the screens of `ItemChecker.html` into an Android app that works without internet. Its libraries are saved inside the app.

**Install on a phone:** copy `ItemChecker.apk` to the phone (USB cable, Drive or Messenger), open it, and allow **Install unknown apps** when asked. To load the list, tap **Import** and choose `items.db` or a CSV/Excel file. Export opens the phone's Share menu.

The phone keeps its own list. It does not sync with the PC: copy `items.db` over and Import it again when the list changes.

**Rebuild after changing `ItemChecker.html`:** the build tools (Java 21 and the Android SDK) are in `%LOCALAPPDATA%\ItemCheckerBuild`. In PowerShell:

```
$B = "$env:LOCALAPPDATA\ItemCheckerBuild"; $env:JAVA_HOME = "$B\jdk"; $env:ANDROID_HOME = "$B\sdk"; $env:GRADLE_USER_HOME = "$B\gradle-home"
cd android-app; npm run sync; cd android; .\gradlew.bat assembleDebug
```

The app file is `android-app\android\app\build\outputs\apk\debug\app-debug.apk`.

## Where the data is saved

Each app keeps its own list on the device:

- **Desktop:** `items.db`, next to `ItemChecker.pyw`.
- **Web and Android app:** on the device. Each browser and phone has its own list. To move items between them, use Export and Import.

The shared online list (Supabase) was turned off on 2026-09-30. The code for it is still there. To turn it back on, put the Project URL and key back into `SUPABASE_URL` and `SUPABASE_ANON_KEY` at the top of `ItemChecker.pyw` and in the `<script>` of `ItemChecker.html`, and run `supabase_setup.sql` in the Supabase SQL Editor.
