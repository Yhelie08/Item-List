"""Item Checker - desktop app.

Stores the item list in a SQLite database (items.db, next to this file) and checks
the Physical Stock of each item against the counts of two verifiers.
Import/export: CSV, Excel (.xlsx), and SQLite (.db) backup/restore.

Shared mode: when SUPABASE_URL and SUPABASE_ANON_KEY are set, the item list lives in
Supabase and items.db is only a local copy. Every change is sent to Supabase, and
changes made by other people (desktop or web) are picked up within a few seconds.
"""
import csv
import datetime
import json
import os
import re
import sqlite3
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
from tkinter import filedialog, messagebox, ttk

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
except ImportError:
    openpyxl = None

if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get('ITEMCHECKER_DB') or os.path.join(APP_DIR, 'items.db')

# Shared online database (Supabase → Project Settings → API). Leave empty to keep the data on this PC only.
SUPABASE_URL = ''
SUPABASE_ANON_KEY = ''
# The web version everyone opens (GitHub Pages). Used by the "Open in Browser" button.
WEB_URL = 'https://yhelie08.github.io/Item-List/'
# Shared mode keeps its local copy here, so the original items.db is never overwritten.
CACHE_PATH = os.environ.get('ITEMCHECKER_CACHE') or os.path.join(APP_DIR, 'items_online_cache.db')
SESSION_PATH = os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'), 'ItemChecker', 'session.json')

HEADERS = ['Item Name', 'SKU', 'ItemRef', 'Warehouse', 'Bin Location', 'Physical Stock QTY',
           'Verifier 1 Count', 'Verifier 2 Count', 'Active']
STATUSES = ['MISMATCH', 'TO COUNT', 'MATCH', 'NO STOCK']

# Older databases still have a committed_stock column; it is no longer used.
SCHEMA = """
CREATE TABLE IF NOT EXISTS items (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  item_name       TEXT    NOT NULL,
  sku             TEXT    NOT NULL UNIQUE COLLATE NOCASE,
  item_ref        TEXT    NOT NULL DEFAULT '',
  warehouse       TEXT    NOT NULL DEFAULT '',
  bin_location    TEXT    NOT NULL DEFAULT '',
  physical_stock  INTEGER NOT NULL DEFAULT 0 CHECK (physical_stock  >= 0),
  verifier1_count INTEGER CHECK (verifier1_count >= 0),
  verifier2_count INTEGER CHECK (verifier2_count >= 0),
  active          INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
  created_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
  updated_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);
CREATE TRIGGER IF NOT EXISTS trg_items_updated AFTER UPDATE ON items
BEGIN
  UPDATE items SET updated_at = datetime('now','localtime') WHERE id = OLD.id;
END;
"""
# Items with stock are counted by two verifiers; an item with 0 stock gets its count typed into Physical Stock.
VIEW = """
CREATE VIEW v_item_check AS
SELECT id, item_name, sku, item_ref, warehouse, bin_location, physical_stock,
       verifier1_count, verifier2_count, active,
       CASE WHEN physical_stock = 0 THEN 'NO STOCK'
            WHEN verifier1_count IS NULL OR verifier2_count IS NULL THEN 'TO COUNT'
            WHEN verifier1_count = physical_stock AND verifier2_count = physical_stock THEN 'MATCH'
            ELSE 'MISMATCH' END AS status,
       updated_at
FROM items;
"""

SAMPLE_ITEMS = [
    ('Nordic Dining Chair', 'HC-CHR-001', 'A-01', 120, 120, 120, 1),
    ('Oak Coffee Table', 'HC-TBL-014', 'A-02', 15, 15, 14, 1),
    ('Rattan Floor Lamp', 'HC-LMP-077', '', 0, None, None, 0),
    ('Linen Throw Pillow', 'HC-PLW-203', 'B-03', 40, None, None, 1),
    ('Ceramic Vase Set', 'HC-DEC-310', 'B-04', 60, 60, None, 1),
]

PRESETS = [
    ('Summary by status', """SELECT status, COUNT(*) AS items, SUM(physical_stock) AS physical
FROM v_item_check
GROUP BY status
ORDER BY CASE status WHEN 'MISMATCH' THEN 1 WHEN 'TO COUNT' THEN 2 WHEN 'MATCH' THEN 3 ELSE 4 END;"""),
    ('Mismatch', """SELECT item_name, sku, bin_location, physical_stock, verifier1_count, verifier2_count
FROM v_item_check
WHERE status = 'MISMATCH'
ORDER BY item_name;"""),
    ('To count (active)', """SELECT item_name, sku, bin_location, physical_stock, verifier1_count, verifier2_count
FROM v_item_check
WHERE active = 1 AND status = 'TO COUNT'
ORDER BY bin_location, item_name;"""),
    ('Inactive items', """SELECT item_name, sku, physical_stock
FROM items
WHERE active = 0
ORDER BY item_name;"""),
    ('Recently updated', """SELECT item_name, sku, physical_stock, verifier1_count, verifier2_count, updated_at
FROM items
ORDER BY updated_at DESC
LIMIT 20;"""),
    ('Schema', "SELECT type, name, sql FROM sqlite_master WHERE name NOT LIKE 'sqlite_%';"),
]


def ensure_schema(conn):
    """Create the tables, and upgrade databases made before the ItemRef / Warehouse / Bin / Verifier columns existed."""
    conn.executescript(SCHEMA)
    have = [c[1] for c in conn.execute('PRAGMA table_info(items)')]
    for col in ('item_ref', 'warehouse', 'bin_location'):
        if col not in have:
            conn.execute(f"ALTER TABLE items ADD COLUMN {col} TEXT NOT NULL DEFAULT ''")
    for col in ('verifier1_count', 'verifier2_count'):
        if col not in have:
            conn.execute(f'ALTER TABLE items ADD COLUMN {col} INTEGER CHECK ({col} >= 0)')
    conn.executescript('DROP VIEW IF EXISTS v_item_check;' + VIEW)


def count_status(physical, v1, v2):
    if physical == 0:
        return 'NO STOCK'
    if v1 is None or v2 is None:
        return 'TO COUNT'
    return 'MATCH' if v1 == physical and v2 == physical else 'MISMATCH'


# ---------------------------------------------------------------- Import parsing

HEADER_ALIASES = {
    'item_name': ['itemname', 'name', 'item', 'productname'],
    'sku': ['sku', 'itemcode', 'skucode'],
    'item_ref': ['itemref', 'cfitemref', 'itemreference', 'reference', 'ref', 'refno', 'referenceno', 'referencenumber'],
    'warehouse': ['warehouse', 'warehousename', 'wh'],
    # Zoho Inventory: CF.Loc is used before Bin Name, which Zoho leaves blank unless bin tracking is on.
    'bin_location': ['binlocation', 'cfloc', 'loc', 'location', 'binname', 'bin'],
    'physical_stock': ['physicalstockqty', 'physicalstock', 'physicalqty', 'physical', 'stockonhand', 'onhand', 'qtyonhand'],
    'verifier1_count': ['verifier1count', 'verifier1'],
    'verifier2_count': ['verifier2count', 'verifier2'],
    'active': ['active', 'isactive', 'status'],
}


def norm_header(h):
    return re.sub(r'[^a-z0-9]', '', str(h if h is not None else '').lower())


def map_headers(row):
    """Column index per field. Earlier aliases win, so a file with both "CF.Loc" and "Bin Name" uses CF.Loc."""
    names = [norm_header(h) for h in row]
    found = {}
    for field, aliases in HEADER_ALIASES.items():
        for a in aliases:
            if a in names:
                found[field] = names.index(a)
                break
    return found


def cell_text(v):
    """Cell value as text, the way it would look in Excel."""
    if v is None:
        return ''
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def parse_qty(v):
    """Whole number >= anything; None if not a whole number. Blank = 0."""
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        return int(v) if v.is_integer() else None
    s = re.sub(r'[,\s]', '', cell_text(v))
    if s == '':
        return 0
    if not re.fullmatch(r'-?\d+(\.0+)?', s):
        return None
    return int(s.split('.')[0])


def check_qty(label, v):
    if v is None:
        raise ValueError(f'{label} must be a whole number.')
    if v < 0:
        raise ValueError(f'{label} cannot be negative.')


def parse_count(v):
    """Verifier count: None when blank (not counted yet), -1 when not a whole number."""
    if cell_text(v) == '':
        return None
    n = parse_qty(v)
    return -1 if n is None else n


def parse_active(v):
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)) and v in (0, 1):
        return int(v)
    s = cell_text(v).lower()
    if s == '':
        return 1
    if s in ('true', 'yes', 'y', '1', 'active'):
        return 1
    if s in ('false', 'no', 'n', '0', 'inactive'):
        return 0
    return None


def read_table_file(path):
    """Return the first sheet of a .csv or .xlsx file as a list of rows."""
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv':
        with open(path, 'rb') as f:
            raw = f.read()
        try:
            text = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = raw.decode('cp1252', errors='replace')
        return list(csv.reader(text.splitlines()))
    if ext in ('.xlsx', '.xlsm'):
        if openpyxl is None:
            raise RuntimeError('Excel support needs openpyxl.  Run:  pip install openpyxl')
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            return [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    raise RuntimeError(f'Unsupported file type "{ext}". Use .csv, .xlsx or a .db backup.')


def split_sql(sql):
    """Split a script into single statements (respecting quotes/triggers)."""
    parts, buf = [], ''
    for chunk in re.split(r'(;)', sql):
        buf += chunk
        if chunk == ';' and sqlite3.complete_statement(buf):
            parts.append(buf)
            buf = ''
    if buf.strip():
        parts.append(buf)
    return [p for p in parts if p.strip().strip(';').strip()]


# ---------------------------------------------------------------- Supabase

FIELDS = ('item_name', 'sku', 'item_ref', 'warehouse', 'bin_location', 'physical_stock',
          'verifier1_count', 'verifier2_count', 'active')


class CloudError(Exception):
    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


def local_time(iso):
    """Supabase timestamp → 'YYYY-MM-DD HH:MM:SS' in local time, like the SQLite defaults."""
    try:
        return datetime.datetime.fromisoformat(iso).astimezone().strftime('%Y-%m-%d %H:%M:%S')
    except (TypeError, ValueError):
        return iso


class Cloud:
    """Minimal Supabase client: password sign-in, and the items / sync_state tables over REST."""
    PAGE = 1000  # Supabase returns at most 1000 rows per request

    def __init__(self, url, key, session_path):
        self.url, self.key, self.session_path = url.rstrip('/'), key, session_path
        self.session = None
        self.lock = threading.Lock()  # the poller thread and the UI share the token

    @property
    def email(self):
        return self.session['email'] if self.session else ''

    def _http(self, method, path, body=None, headers=None, token=None):
        # New-style publishable keys (sb_publishable_...) are not JWTs, so only a user token goes in Authorization.
        h = {'apikey': self.key, 'Content-Type': 'application/json'}
        if token:
            h['Authorization'] = f'Bearer {token}'
        h.update(headers or {})
        data = json.dumps(body).encode('utf-8') if body is not None else None
        req = urllib.request.Request(self.url + path, data=data, method=method, headers=h)
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                raw = r.read()
            return json.loads(raw) if raw else None
        except urllib.error.HTTPError as ex:
            msg = ex.read().decode('utf-8', 'replace')
            try:
                j = json.loads(msg)
                msg = j.get('message') or j.get('msg') or j.get('error_description') or msg
            except ValueError:
                pass
            raise CloudError(msg, ex.code)
        except (urllib.error.URLError, TimeoutError, OSError) as ex:
            raise CloudError(f'Cannot reach the online database ({getattr(ex, "reason", ex)}). Check the internet connection.')

    def _auth(self, grant, body):
        s = self._http('POST', f'/auth/v1/token?grant_type={grant}', body)
        self.session = {'access_token': s['access_token'], 'refresh_token': s['refresh_token'],
                        'email': (s.get('user') or {}).get('email', ''),
                        'expires_at': time.time() + int(s.get('expires_in', 3600)) - 60}
        try:
            os.makedirs(os.path.dirname(self.session_path), exist_ok=True)
            with open(self.session_path, 'w', encoding='utf-8') as f:
                json.dump({'refresh_token': s['refresh_token'], 'email': self.session['email']}, f)
        except OSError:
            pass

    def sign_in(self, email, password):
        try:
            self._auth('password', {'email': email.strip(), 'password': password})
        except CloudError as ex:
            if ex.status in (400, 401):
                raise CloudError('Wrong email or password.', ex.status)
            raise

    def resume(self):
        """Sign in again with the saved session. Returns False if the user has to log in."""
        try:
            with open(self.session_path, encoding='utf-8') as f:
                token = json.load(f)['refresh_token']
            self._auth('refresh_token', {'refresh_token': token})
            return True
        except (OSError, ValueError, KeyError, CloudError):
            return False

    def sign_out(self):
        if self.session:
            try:
                self._http('POST', '/auth/v1/logout', {}, token=self.session['access_token'])
            except CloudError:
                pass
        self.session = None
        try:
            os.remove(self.session_path)
        except OSError:
            pass

    def api(self, method, path, body=None, headers=None):
        with self.lock:
            if time.time() > self.session['expires_at']:
                self._auth('refresh_token', {'refresh_token': self.session['refresh_token']})
            token = self.session['access_token']
        return self._http(method, '/rest/v1/' + path, body, headers, token)

    def fetch_items(self):
        rows, offset = [], 0
        while True:
            page = self.api('GET', f'items?select=id,{",".join(FIELDS)},created_at,updated_at'
                                   f'&order=id&limit={self.PAGE}&offset={offset}')
            for r in page:
                for f in ('item_ref', 'warehouse', 'bin_location'):
                    if r.get(f) is None:
                        r[f] = ''
                for f in ('verifier1_count', 'verifier2_count'):
                    r.setdefault(f, None)
            rows += page
            if len(page) < self.PAGE:
                return rows
            offset += self.PAGE

    def version(self):
        """Counter bumped by a trigger on every change to items."""
        r = self.api('GET', 'sync_state?select=version&id=eq.1')
        if not r:
            raise CloudError('The sync_state table is empty. Run supabase_setup.sql in the Supabase SQL Editor.')
        return r[0]['version']

    def push(self, deletes, updates, inserts):
        for i in range(0, len(deletes), 200):
            self.api('DELETE', f'items?id=in.({",".join(map(str, deletes[i:i + 200]))})')
        for i in range(0, len(updates), 500):
            self.api('POST', 'items?on_conflict=id', updates[i:i + 500],
                     {'Prefer': 'resolution=merge-duplicates,return=minimal'})
        for i in range(0, len(inserts), 500):
            self.api('POST', 'items', inserts[i:i + 500], {'Prefer': 'return=minimal'})


# ---------------------------------------------------------------- Data layer

class Store:
    def __init__(self, path, seed=True):
        self.path = path
        self.snapshot = {}  # shared mode: id -> FIELDS as last loaded from Supabase
        is_new = not os.path.exists(path)
        self.conn = sqlite3.connect(path, isolation_level=None)
        self.conn.row_factory = sqlite3.Row
        ensure_schema(self.conn)
        if is_new and seed:
            self.conn.execute('BEGIN')
            self.conn.executemany(
                'INSERT INTO items (item_name, sku, bin_location, physical_stock, verifier1_count, verifier2_count, active) '
                'VALUES (?,?,?,?,?,?,?)', SAMPLE_ITEMS)
            self.conn.execute('COMMIT')
            self.sample_flag = True

    # PRAGMA user_version = 1 marks "sample data still loaded" (shows the banner).
    @property
    def sample_flag(self):
        return self.conn.execute('PRAGMA user_version').fetchone()[0] == 1

    @sample_flag.setter
    def sample_flag(self, on):
        self.conn.execute(f'PRAGMA user_version = {1 if on else 0}')

    def close(self):
        self.conn.close()

    # ---------- shared mode ----------
    def load_remote(self, rows):
        """Replace the local copy with the rows from Supabase."""
        self.conn.execute('BEGIN')
        try:
            self.conn.execute('DELETE FROM items')
            self.conn.executemany(
                f'INSERT INTO items (id, {", ".join(FIELDS)}, created_at, updated_at) '
                f'VALUES ({",".join("?" * (len(FIELDS) + 3))})',
                [(r['id'], *(r[f] for f in FIELDS), local_time(r['created_at']), local_time(r['updated_at'])) for r in rows])
            self.conn.execute('COMMIT')
        except Exception:
            self.conn.execute('ROLLBACK')
            raise
        self.sample_flag = False
        self.snapshot = {r['id']: tuple(r[f] for f in FIELDS) for r in rows}

    def changes(self):
        """What changed locally since load_remote(): (deleted ids, updated rows, new rows)."""
        updates, inserts, seen = [], [], set()
        for r in self.conn.execute(f'SELECT id, {", ".join(FIELDS)} FROM items'):
            row = {f: r[f] for f in FIELDS}
            old = self.snapshot.get(r['id'])
            if old is None:
                inserts.append(row)
            else:
                seen.add(r['id'])
                if old != tuple(row.values()):
                    updates.append({'id': r['id'], **row})
        return [i for i in self.snapshot if i not in seen], updates, inserts

    def list_items(self, search='', active='all', status='all', sort_col='item_name', sort_dir='asc'):
        where, params = [], []
        if search:
            like = '%' + re.sub(r'([\\%_])', r'\\\1', search) + '%'
            where.append("(item_name LIKE ? ESCAPE '\\' OR sku LIKE ? ESCAPE '\\' OR item_ref LIKE ? ESCAPE '\\' "
                         "OR warehouse LIKE ? ESCAPE '\\' OR bin_location LIKE ? ESCAPE '\\')")
            params += [like] * 5
        if active in ('0', '1'):
            where.append('active = ?')
            params.append(int(active))
        if status in STATUSES:
            where.append('status = ?')
            params.append(status)
        cols = ['item_name', 'sku', 'item_ref', 'warehouse', 'bin_location', 'physical_stock', 'status', 'active', 'updated_at']
        col = sort_col if sort_col in cols else 'item_name'
        collate = ' COLLATE NOCASE' if col in ('item_name', 'sku', 'item_ref', 'warehouse', 'bin_location', 'status') else ''
        direction = 'DESC' if sort_dir == 'desc' else 'ASC'
        sql = (f"SELECT * FROM v_item_check {'WHERE ' + ' AND '.join(where) if where else ''} "
               f"ORDER BY {col}{collate} {direction}, item_name COLLATE NOCASE")
        return self.conn.execute(sql, params).fetchall()

    def totals(self):
        return self.conn.execute("""
            SELECT COUNT(*) AS total,
                   COALESCE(SUM(active), 0) AS active,
                   COALESCE(SUM(physical_stock), 0) AS physical,
                   COALESCE(SUM(status = 'TO COUNT'), 0) AS to_count,
                   COALESCE(SUM(status = 'MISMATCH'), 0) AS mismatch
            FROM v_item_check""").fetchone()

    def get(self, item_id):
        return self.conn.execute('SELECT * FROM items WHERE id = ?', (item_id,)).fetchone()

    def get_check(self, item_id):
        return self.conn.execute('SELECT * FROM v_item_check WHERE id = ?', (item_id,)).fetchone()

    def suggestions(self, col):
        """Warehouses or bin locations already in use, for the Add/Edit suggestions."""
        return [r[0] for r in self.conn.execute(
            f"SELECT DISTINCT {col} FROM items WHERE {col} <> '' ORDER BY {col} COLLATE NOCASE")]

    def find_sku(self, sku, except_id=None):
        return self.conn.execute('SELECT id, item_name, sku FROM items WHERE sku = ? AND id IS NOT ?',
                                 (sku, except_id)).fetchone()

    def add_item(self, name, sku, item_ref, warehouse, bin_location, physical, active):
        """Validate and insert a new item. Raises ValueError with a user message."""
        name, sku, item_ref, warehouse, bin_location = (
            (v or '').strip() for v in (name, sku, item_ref, warehouse, bin_location))
        if not name:
            raise ValueError('Item Name is required.')
        if not sku:
            raise ValueError('SKU is required.')
        check_qty('Physical Stock QTY', physical)
        dup = self.find_sku(sku)
        if dup:
            raise ValueError(f'SKU "{dup["sku"]}" already exists ({dup["item_name"]}). SKUs are not case-sensitive.')
        self.conn.execute('INSERT INTO items (item_name, sku, item_ref, warehouse, bin_location, physical_stock, active) '
                          'VALUES (?,?,?,?,?,?,?)', (name, sku, item_ref, warehouse, bin_location, physical, int(bool(active))))

    def update_item(self, item_id, item_ref, warehouse, bin_location, active, physical=None, counts=None):
        """Item Name and SKU of a saved item never change. Pass physical (item with 0 stock)
        or counts=(verifier 1, verifier 2) (item with stock; None = not counted yet)."""
        item_ref, warehouse, bin_location = ((v or '').strip() for v in (item_ref, warehouse, bin_location))
        sets, params = ['item_ref=?', 'warehouse=?', 'bin_location=?', 'active=?'], [item_ref, warehouse, bin_location, int(bool(active))]
        if physical is not None or counts is None:
            check_qty('Physical Stock QTY', physical)
            sets.append('physical_stock=?')
            params.append(physical)
        else:
            for label, v in zip(('Verifier 1 (Count)', 'Verifier 2 (Count)'), counts):
                if v is not None:
                    check_qty(label, v)
            sets += ['verifier1_count=?', 'verifier2_count=?']
            params += list(counts)
        self.conn.execute(f'UPDATE items SET {", ".join(sets)} WHERE id=?', (*params, item_id))

    def update_verifier_count(self, item_id, field, val):
        """Update a single verifier count (verifier1_count or verifier2_count). None means not counted yet."""
        if field not in ('verifier1_count', 'verifier2_count'):
            raise ValueError(f'Invalid field: {field}')
        if val is not None:
            check_qty('Verifier Count', val)
        self.conn.execute(f'UPDATE items SET {field} = ? WHERE id = ?', (val, item_id))

    def toggle_active(self, item_id):
        self.conn.execute('UPDATE items SET active = 1 - active WHERE id = ?', (item_id,))

    def delete(self, item_id):
        self.conn.execute('DELETE FROM items WHERE id = ?', (item_id,))

    def clear_all(self):
        self.conn.execute('BEGIN')
        self.conn.execute('DELETE FROM items')
        self.conn.execute("DELETE FROM sqlite_sequence WHERE name = 'items'")
        self.conn.execute('COMMIT')
        self.sample_flag = False

    def import_rows(self, rows):
        """UPSERT rows (first sheet incl. header). Returns (added, updated, skipped[(row_no, sku, reason)])."""
        header_idx, cols = -1, None
        for i, row in enumerate(rows[:10]):
            m = map_headers(row)
            if 'item_name' in m and 'sku' in m:
                header_idx, cols = i, m
                break
        if header_idx < 0:
            raise ValueError('Could not find the header row. The file needs at least the columns "Item Name" and "SKU".')

        def get(row, field):
            i = cols.get(field)
            return row[i] if i is not None and i < len(row) else None

        # Files without an ItemRef, Warehouse, Bin Location or Verifier column leave the saved values alone.
        set_ref = ''.join(f'{c} = excluded.{c}, ' for c in
                          ('item_ref', 'warehouse', 'bin_location', 'verifier1_count', 'verifier2_count') if c in cols)
        added = updated = 0
        skipped = []
        self.conn.execute('BEGIN')
        try:
            for i in range(header_idx + 1, len(rows)):
                row, row_no = rows[i], i + 1
                if not row or all(cell_text(c) == '' for c in row):
                    continue
                name, sku = cell_text(get(row, 'item_name')), cell_text(get(row, 'sku'))
                phys = parse_qty(get(row, 'physical_stock'))
                v1, v2 = parse_count(get(row, 'verifier1_count')), parse_count(get(row, 'verifier2_count'))
                active = parse_active(get(row, 'active'))
                why = None
                if not name:
                    why = 'Item Name is required'
                elif not sku:
                    why = 'SKU is required'
                elif phys is None:
                    why = f'Physical Stock QTY "{cell_text(get(row, "physical_stock"))}" is not a whole number'
                elif phys < 0:
                    why = 'Physical Stock QTY cannot be negative'
                elif v1 is not None and v1 < 0:
                    why = f'Verifier 1 Count "{cell_text(get(row, "verifier1_count"))}" is not a whole number, 0 or more'
                elif v2 is not None and v2 < 0:
                    why = f'Verifier 2 Count "{cell_text(get(row, "verifier2_count"))}" is not a whole number, 0 or more'
                elif active is None:
                    why = f'Active "{cell_text(get(row, "active"))}" is not TRUE/FALSE, YES/NO, 1/0 or Active'
                if why:
                    skipped.append((row_no, sku, why))
                    continue
                exists = self.find_sku(sku) is not None
                self.conn.execute(f"""
                    INSERT INTO items (item_name, sku, item_ref, warehouse, bin_location, physical_stock,
                                       verifier1_count, verifier2_count, active)
                    VALUES (?,?,?,?,?,?,?,?,?)
                    ON CONFLICT(sku) DO UPDATE SET item_name = excluded.item_name, {set_ref}
                      physical_stock = excluded.physical_stock, active = excluded.active""",
                    (name, sku, cell_text(get(row, 'item_ref')), cell_text(get(row, 'warehouse')),
                     cell_text(get(row, 'bin_location')), phys, v1, v2, active))
                if exists:
                    updated += 1
                else:
                    added += 1
            self.conn.execute('COMMIT')
        except Exception:
            self.conn.execute('ROLLBACK')
            raise
        return added, updated, skipped

    def export_rows(self):
        return self.conn.execute('SELECT item_name, sku, item_ref, warehouse, bin_location, physical_stock, '
                                 'verifier1_count, verifier2_count, active, status FROM v_item_check '
                                 'ORDER BY item_name COLLATE NOCASE, sku').fetchall()

    def backup_to(self, path):
        if os.path.abspath(path) == os.path.abspath(self.path):
            raise ValueError('Choose a different file name than the live database.')
        if os.path.exists(path):
            os.remove(path)
        dst = sqlite3.connect(path)
        try:
            self.conn.backup(dst)
        finally:
            dst.close()

    def restore_from(self, path):
        """Replace all data with a .db backup. The backup file itself is not modified."""
        src = sqlite3.connect(path)
        mem = sqlite3.connect(':memory:')
        try:
            try:
                src.backup(mem)
                if not mem.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='items'").fetchone():
                    raise ValueError('This .db file does not contain an items table.')
                ensure_schema(mem)
                mem.execute('SELECT id, item_name, sku, item_ref, warehouse, status FROM v_item_check LIMIT 1').fetchall()
            except sqlite3.DatabaseError as ex:
                raise ValueError(f'This file is not a valid Item Checker database ({ex}).')
            mem.execute('PRAGMA user_version = 0')
            mem.backup(self.conn)
        finally:
            src.close()
            mem.close()
        return self.conn.execute('SELECT COUNT(*) FROM items').fetchone()[0]

    def run_sql(self, sql):
        """Run one or more statements. Returns (columns, rows, rows_changed, any_change)."""
        before = self.conn.total_changes
        columns, rows, changed = None, None, 0
        for stmt in split_sql(sql):
            cur = self.conn.execute(stmt)
            if cur.description:
                columns = [d[0] for d in cur.description]
                rows = [tuple(r) for r in cur.fetchall()]
            elif cur.rowcount > 0:
                changed += cur.rowcount
        return columns, rows, changed, self.conn.total_changes > before


# ---------------------------------------------------------------- File writers

def style_header(ws):
    fill = PatternFill('solid', fgColor='0F766E')
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = fill
    for col, w in zip('ABCDEFGHIJ', (34, 16, 16, 16, 14, 18, 16, 16, 9, 14)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'


# "Count Result" is for reading only; Import ignores it.
EXPORT_HEADERS = HEADERS + ['Count Result']


def export_values(r, active):
    return [r['item_name'], r['sku'], r['item_ref'], r['warehouse'], r['bin_location'], r['physical_stock'],
            '' if r['verifier1_count'] is None else r['verifier1_count'],
            '' if r['verifier2_count'] is None else r['verifier2_count'], active, r['status']]


def write_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Items'
    ws.append(EXPORT_HEADERS)
    for r in rows:
        ws.append(export_values(r, bool(r['active'])))
    style_header(ws)
    wb.save(path)


def write_csv(path, header, rows):
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def write_template(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Items'
    ws.append(HEADERS)
    style_header(ws)
    help_ws = wb.create_sheet('Instructions')
    for line in [
        ['How to fill the Items sheet'],
        ['Item Name and SKU are required.'],
        ['SKU must be unique. It is not case-sensitive (HC-001 = hc-001). An existing SKU is updated on import.'],
        ['ItemRef is optional. If the file has no ItemRef column, existing ItemRefs are kept.'],
        ['Warehouse and Bin Location are optional. If the file has no such column, the saved values are kept.'],
        ['Physical Stock QTY: whole number, 0 or more. Blank = 0. Commas are fine (1,200).'],
        ['Verifier 1 Count and Verifier 2 Count: optional whole numbers. Blank = not counted yet. '
         'No such column = saved counts are kept.'],
        ['Active: TRUE/FALSE, YES/NO, 1/0, or Active/Inactive. Blank = TRUE.'],
        [],
        ['Example'],
        HEADERS,
        ['Nordic Dining Chair', 'HC-CHR-001', 'REF-1001', 'WH1 BANAWE', 'A-01', 120, 120, 120, 'TRUE'],
        ['Oak Coffee Table', 'HC-TBL-014', 'REF-1002', 'WH1 BANAWE', 'A-02', 15, '', '', 'TRUE'],
        ['Rattan Floor Lamp', 'HC-LMP-077', '', '', '', 0, '', '', 'FALSE'],
    ]:
        help_ws.append(line)
    help_ws['A1'].font = Font(bold=True, size=12)
    help_ws.column_dimensions['A'].width = 34
    wb.save(path)


# ---------------------------------------------------------------- UI

TEAL, TEAL_DARK, BG, CARD, MUTED, RED = '#0f766e', '#115e59', '#f1f5f9', '#ffffff', '#64748b', '#dc2626'
AMBER = '#b45309'
STATUS_COLORS = {'MISMATCH': RED, 'TO COUNT': AMBER, 'MATCH': '#047857', 'NO STOCK': MUTED}
FONT = ('Segoe UI', 10)


def today():
    return datetime.date.today().isoformat()


def flow(frame, widgets, gap=6, align='left'):
    """Lay widgets out left to right, wrapping onto a new row when the frame gets too narrow."""
    def relayout(e=None):
        width = frame.winfo_width()
        if width <= 1:
            return
        rows = []
        cur_row = []
        cur_w = 0
        for w in widgets:
            ww = w.winfo_reqwidth()
            if cur_row and cur_w + gap + ww > width:
                rows.append((cur_row, cur_w))
                cur_row = [w]
                cur_w = ww
            else:
                cur_row.append(w)
                cur_w += (gap if cur_row else 0) + ww
        if cur_row:
            rows.append((cur_row, cur_w))

        y = 0
        for row_widgets, row_w in rows:
            row_h = max(w.winfo_reqheight() for w in row_widgets)
            x = max(0, (width - row_w) // 2) if align == 'center' else 0
            for w in row_widgets:
                w.place(x=x, y=y)
                x += w.winfo_reqwidth() + gap
            y += row_h + gap
        total_h = max(0, y - gap)
        if int(frame.cget('height')) != total_h:
            frame.configure(height=total_h)
    frame.bind('<Configure>', relayout, add='+')
    for w in widgets:  # e.g. a label whose text got longer
        w.bind('<Configure>', relayout, add='+')
    frame.after_idle(relayout)


class ItemCardFrame(tk.Frame):
    """Every field of one item, with direct verifier counts entry, toggle active, and delete."""

    def __init__(self, parent, app, show_close=False, on_deleted=None):
        super().__init__(parent, bg=CARD, padx=18, pady=14,
                         highlightbackground='#cbd5e1', highlightthickness=1)
        self.app = app
        self.show_close = show_close
        self.on_deleted = on_deleted
        self.item_id = None
        self.unlocked = {'verifier1_count': False, 'verifier2_count': False}
        self.stat_boxes = {}
        self.v_entries = {}
        self.columnconfigure(0, weight=1)

        self.l_name = tk.Label(self, bg=CARD, fg='#0f172a', font=('Segoe UI Semibold', 13), anchor='w', justify='left')
        self.l_name.grid(row=0, column=0, sticky='we')
        ttk.Separator(self).grid(row=1, column=0, sticky='we', pady=(10, 4))

        def field(title, row, font):
            tk.Label(self, text=title.upper(), bg=CARD, fg=MUTED, font=('Segoe UI Semibold', 8)).grid(
                row=row, column=0, sticky='w', pady=(8, 0))
            v = tk.Label(self, bg=CARD, fg='#0f172a', font=font, anchor='w', justify='left')
            v.grid(row=row + 1, column=0, sticky='we')
            return v

        self.l_sku = field('SKU', 2, ('Consolas', 11))
        self.l_ref = field('ItemRef', 4, FONT)
        self.l_wh = field('Warehouse', 6, FONT)
        self.l_bin = field('Bin Location', 8, FONT)

        stats = tk.Frame(self, bg=CARD)
        stats.grid(row=10, column=0, sticky='we', pady=(12, 0))
        for i, (key, title) in enumerate((('verifier1_count', 'Verifier 1'),
                                          ('verifier2_count', 'Verifier 2'))):
            stats.columnconfigure(i, weight=1, uniform='s')
            box = tk.Frame(stats, bg=BG, padx=10, pady=8)
            box.grid(row=0, column=i, sticky='nsew', padx=(0 if i == 0 else 8, 0))
            tk.Label(box, text=title, bg=BG, fg=MUTED, font=('Segoe UI', 9)).pack(anchor='w')
            content = tk.Frame(box, bg=BG)
            content.pack(fill='x', expand=True, pady=(4, 0))
            self.stat_boxes[key] = content

        res = tk.Frame(self, bg=BG, padx=10, pady=6)
        res.grid(row=11, column=0, sticky='we', pady=(8, 0))
        tk.Label(res, text='Count result', bg=BG, font=('Segoe UI Semibold', 10)).pack(side='left')
        self.l_status = tk.Label(res, bg=BG, font=('Segoe UI Semibold', 10))
        self.l_status.pack(side='right')

        act = tk.Frame(self, bg=BG, padx=10, pady=6)
        act.grid(row=12, column=0, sticky='we', pady=(8, 0))
        tk.Label(act, text='Active', bg=BG, font=('Segoe UI Semibold', 10)).pack(side='left')
        self.b_toggle = ttk.Button(act, command=self.toggle)
        self.b_toggle.pack(side='right')
        self.l_active = tk.Label(act, bg=BG)
        self.l_active.pack(side='right', padx=(0, 8))

        btns = tk.Frame(self, bg=CARD)
        btns.grid(row=13, column=0, sticky='we', pady=(14, 0))
        ttk.Button(btns, text='Delete', command=self.delete).pack(side='left')
        if self.show_close:
            self.b_close = ttk.Button(btns, text='Close', command=self.close)
            self.b_close.pack(side='right')
        else:
            self.b_close = None

        self.bind('<Configure>', self.rewrap)

    def rewrap(self, e=None):
        width = max(self.winfo_width() - 40, 200)
        for lbl in (self.l_name, self.l_sku, self.l_ref, self.l_wh, self.l_bin):
            lbl.configure(wraplength=width)

    def load_item(self, item_id):
        self.item_id = item_id
        self.unlocked['verifier1_count'] = False
        self.unlocked['verifier2_count'] = False
        ok = self.fill()
        if ok:
            def set_initial_focus():
                if 'verifier1_count' in self.v_entries:
                    self.v_entries['verifier1_count'].focus_set()
                    self.v_entries['verifier1_count'].selection_range(0, 'end')
                elif 'verifier2_count' in self.v_entries:
                    self.v_entries['verifier2_count'].focus_set()
                    self.v_entries['verifier2_count'].selection_range(0, 'end')
            self.after(50, set_initial_focus)
        return ok

    def render_verifiers(self, r):
        for key in ('verifier1_count', 'verifier2_count'):
            box = self.stat_boxes[key]
            for child in box.winfo_children():
                child.destroy()
            val = r[key]
            in_edit = (val is None or self.unlocked[key])
            if in_edit:
                entry = ttk.Entry(box, font=('Segoe UI', 10), width=8)
                if val is not None:
                    entry.insert(0, str(val))
                entry.pack(side='left', fill='x', expand=True)
                self.v_entries[key] = entry
                entry.bind('<Return>', lambda e, k=key: self.on_entry_enter(k))
                entry.bind('<KP_Enter>', lambda e, k=key: self.on_entry_enter(k))

                btn_save = ttk.Button(box, text='Save', width=5,
                                      command=lambda k=key: self.save_field(k))
                btn_save.pack(side='right', padx=(4, 0))
            else:
                self.v_entries.pop(key, None)
                off = (val != r['physical_stock'])
                lbl = tk.Label(box, text=f'{val:,}', bg=BG,
                               font=('Segoe UI Semibold', 16), fg=RED if off else '#0f172a')
                lbl.pack(side='left')
                btn_edit = ttk.Button(box, text='Edit', width=5,
                                      command=lambda k=key: self.unlock_field(k))
                btn_edit.pack(side='right')

    def unlock_field(self, key):
        self.unlocked[key] = True
        r = self.app.store.get_check(self.item_id)
        if r:
            self.render_verifiers(r)
            if key in self.v_entries:
                self.v_entries[key].focus_set()
                self.v_entries[key].selection_range(0, 'end')

    def save_field(self, key):
        entry = self.v_entries.get(key)
        if not entry:
            return True
        raw = entry.get().strip()
        if raw == '':
            val = None
        else:
            try:
                clean = raw.replace(',', '').replace(' ', '')
                val = int(clean)
                if val < 0:
                    raise ValueError()
            except ValueError:
                messagebox.showerror('Invalid Count', 'Verifier count must be a whole number, 0 or more.', parent=self)
                entry.focus_set()
                entry.selection_range(0, 'end')
                return False
        try:
            self.app.store.update_verifier_count(self.item_id, key, val)
        except Exception as ex:
            messagebox.showerror('Save Error', str(ex), parent=self)
            return False
        self.unlocked[key] = False
        self.app.saved()
        self.fill()
        return True

    def on_entry_enter(self, key):
        if self.save_field(key):
            other_key = 'verifier2_count' if key == 'verifier1_count' else 'verifier1_count'
            if other_key in self.v_entries:
                self.v_entries[other_key].focus_set()
                self.v_entries[other_key].selection_range(0, 'end')

    def fill(self):
        if not self.item_id:
            return False
        r = self.app.store.get_check(self.item_id)
        if not r:
            return False
        self.l_name.config(text=r['item_name'])
        self.l_sku.config(text=r['sku'])
        self.l_ref.config(text=r['item_ref'] or '—')
        self.l_wh.config(text=r['warehouse'] or '—')
        self.l_bin.config(text=r['bin_location'] or '—')
        self.render_verifiers(r)
        self.l_status.config(text=r['status'], fg=STATUS_COLORS[r['status']])
        self.l_active.config(text='Yes' if r['active'] else 'No', fg=TEAL if r['active'] else MUTED)
        self.b_toggle.config(text='Mark inactive' if r['active'] else 'Mark active')
        return True

    def toggle(self):
        if not self.item_id:
            return
        self.app.store.toggle_active(self.item_id)
        self.app.saved()
        self.fill()

    def delete(self):
        if not self.item_id:
            return
        r = self.app.store.get(self.item_id)
        if r and messagebox.askyesno('Delete item?', f'{r["item_name"]} ({r["sku"]}) will be permanently removed.',
                                     icon='warning', parent=self):
            deleted_id = self.item_id
            self.item_id = None
            self.app.store.delete(deleted_id)
            self.app.saved()
            self.app.set_status(f'Deleted {r["sku"]}')
            if self.on_deleted:
                self.on_deleted(deleted_id)
            else:
                self.close()

    def close(self):
        master = self.winfo_toplevel()
        if master != self.app:
            master.destroy()


class ItemCard(tk.Toplevel):
    """Every field of one item, opened from the item list as popup."""

    def __init__(self, app, item_id):
        super().__init__(app)
        self.app, self.item_id = app, item_id
        self.title('Item')
        self.configure(bg=CARD)
        self.transient(app)
        self.minsize(320, 0)
        self.card = ItemCardFrame(self, app, show_close=True, on_deleted=lambda _: self.destroy())
        self.card.pack(fill='both', expand=True)
        if not self.card.load_item(item_id):
            self.destroy()
            return
        self.bind('<Escape>', lambda e: self.destroy())
        self.update_idletasks()
        w = max(self.winfo_reqwidth(), 440)
        x = app.winfo_rootx() + (app.winfo_width() - w) // 2
        y = app.winfo_rooty() + (app.winfo_height() - self.winfo_reqheight()) // 3
        self.geometry(f'{w}x{self.winfo_reqheight()}+{max(x, 0)}+{max(y, 0)}')
        self.grab_set()


class ItemDialog(tk.Toplevel):
    """Add a new item, or edit a saved one. A saved item's Item Name and SKU are read-only.
    With 0 Physical Stock the count goes straight into Physical Stock; with stock,
    Physical Stock is locked and the two verifiers enter their counts."""

    def __init__(self, app, item=None, reopen_card=False):
        super().__init__(app)
        self.app, self.item, self.reopen_card = app, item, reopen_card
        self.locked = bool(item) and item['physical_stock'] > 0
        self.title('Edit Item' if item else 'Add Item')
        self.configure(bg=CARD, padx=18, pady=14)
        self.resizable(False, False)
        self.transient(app)

        def val(key, blank=''):
            return blank if not item or item[key] is None else str(item[key])

        self.v_name = tk.StringVar(value=val('item_name'))
        self.v_sku = tk.StringVar(value=val('sku'))
        self.v_ref = tk.StringVar(value=val('item_ref'))
        self.v_wh = tk.StringVar(value=val('warehouse'))
        self.v_bin = tk.StringVar(value=val('bin_location'))
        self.v_phys = tk.StringVar(value=val('physical_stock', '0'))
        self.v_c1 = tk.StringVar(value=val('verifier1_count'))
        self.v_c2 = tk.StringVar(value=val('verifier2_count'))
        self.v_active = tk.BooleanVar(value=bool(item['active']) if item else True)

        def label(text, r, c=0):
            tk.Label(self, text=text, bg=CARD, font=('Segoe UI Semibold', 10)).grid(row=r, column=c, sticky='w', pady=(6, 2))

        ro = ['readonly'] if item else []
        label('Item Name' if item else 'Item Name *', 0)
        e_name = ttk.Entry(self, textvariable=self.v_name, width=44, font=FONT)
        e_name.state(ro)
        e_name.grid(row=1, column=0, columnspan=2, sticky='we')
        label('SKU' if item else 'SKU *', 2)
        e_sku = ttk.Entry(self, textvariable=self.v_sku, width=44, font=('Consolas', 10))
        e_sku.state(ro)
        e_sku.grid(row=3, column=0, columnspan=2, sticky='we')
        label('ItemRef', 4)
        ttk.Entry(self, textvariable=self.v_ref, width=44, font=FONT).grid(row=5, column=0, columnspan=2, sticky='we')
        label('Warehouse', 6)
        ttk.Combobox(self, textvariable=self.v_wh, values=app.store.suggestions('warehouse'), width=42, font=FONT).grid(
            row=7, column=0, columnspan=2, sticky='we')
        label('Bin Location', 8)
        ttk.Combobox(self, textvariable=self.v_bin, values=app.store.suggestions('bin_location'), width=42, font=FONT).grid(
            row=9, column=0, columnspan=2, sticky='we')
        label('Physical Stock QTY', 10)
        e_phys = ttk.Spinbox(self, from_=0, to=10**9, textvariable=self.v_phys, width=18, font=FONT)
        e_phys.grid(row=11, column=0, sticky='w', padx=(0, 8))
        first = e_phys if item else e_name
        if self.locked:
            e_phys.state(['disabled'])  # 'readonly' would still let the arrows change it
            label('Verifier 1 (Count)', 12, 0)
            label('Verifier 2 (Count)', 12, 1)
            e_c1 = ttk.Entry(self, textvariable=self.v_c1, width=20, font=FONT)
            e_c1.grid(row=13, column=0, sticky='w', padx=(0, 8))
            ttk.Entry(self, textvariable=self.v_c2, width=20, font=FONT).grid(row=13, column=1, sticky='w')
            first = e_c1
        ttk.Checkbutton(self, text='Active', variable=self.v_active).grid(row=14, column=0, sticky='w', pady=(10, 0))

        self.preview = tk.Label(self, bg=CARD, fg=MUTED, font=FONT, anchor='w')
        self.preview.grid(row=15, column=0, columnspan=2, sticky='we', pady=(8, 0))
        self.error = tk.Label(self, bg='#fef2f2', fg='#b91c1c', font=FONT, anchor='w', justify='left', wraplength=380)
        btns = tk.Frame(self, bg=CARD)
        btns.grid(row=17, column=0, columnspan=2, sticky='e', pady=(14, 0))
        ttk.Button(btns, text='Cancel', command=self.destroy).pack(side='right')
        ttk.Button(btns, text='Save', style='Accent.TButton', command=self.save).pack(side='right', padx=(0, 6))

        for v in (self.v_phys, self.v_c1, self.v_c2):
            v.trace_add('write', lambda *_: self.update_preview())
        self.update_preview()
        self.bind('<Return>', lambda e: self.save())
        self.bind('<Escape>', lambda e: self.destroy())

        self.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - self.winfo_width()) // 2
        y = app.winfo_rooty() + (app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f'+{max(x, 0)}+{max(y, 0)}')
        self.grab_set()
        first.focus_set()
        if first is e_phys:
            e_phys.selection_range(0, 'end')

    def counts(self):
        return tuple(parse_count(v.get()) if self.locked else None for v in (self.v_c1, self.v_c2))

    def update_preview(self):
        p, (c1, c2) = parse_qty(self.v_phys.get()), self.counts()
        if p is None or p < 0 or -1 in (c1, c2):
            self.preview.config(text='')
        else:
            s = count_status(p, c1, c2)
            self.preview.config(text=f'Count result: {s}', fg=STATUS_COLORS[s])

    def save(self):
        s = self.app.store
        try:
            if not self.item:
                s.add_item(self.v_name.get(), self.v_sku.get(), self.v_ref.get(), self.v_wh.get(), self.v_bin.get(),
                           parse_qty(self.v_phys.get()), self.v_active.get())
            elif self.locked:
                c = self.counts()
                for label, v in zip(('Verifier 1 (Count)', 'Verifier 2 (Count)'), c):
                    if v == -1:
                        raise ValueError(f'{label} must be a whole number.')
                s.update_item(self.item['id'], self.v_ref.get(), self.v_wh.get(), self.v_bin.get(), self.v_active.get(),
                              counts=c)
            else:
                s.update_item(self.item['id'], self.v_ref.get(), self.v_wh.get(), self.v_bin.get(), self.v_active.get(),
                              physical=parse_qty(self.v_phys.get()))
        except ValueError as ex:
            self.error.config(text=str(ex), padx=8, pady=6)
            self.error.grid(row=16, column=0, columnspan=2, sticky='we', pady=(8, 0))
            return
        sku = self.v_sku.get().strip()
        self.destroy()
        self.app.saved()
        self.app.set_status(f'{"Updated" if self.item else "Added"} {sku}')
        if self.reopen_card:
            ItemCard(self.app, self.item['id'])


class LoginDialog(tk.Toplevel):
    """Sign in to the shared online list. self.ok is True after a successful sign-in."""

    def __init__(self, app):
        super().__init__(app)
        self.app, self.ok = app, False
        self.title('Sign in - Item Checker')
        self.configure(bg=CARD, padx=18, pady=14)
        self.resizable(False, False)
        self.transient(app)
        tk.Label(self, text='Sign in to the shared item list', bg=CARD, font=('Segoe UI Semibold', 12)).grid(
            row=0, column=0, sticky='w')
        self.v_email = tk.StringVar(value=self.app.cloud.email)
        self.v_pw = tk.StringVar()
        tk.Label(self, text='Email', bg=CARD, font=('Segoe UI Semibold', 10)).grid(row=1, column=0, sticky='w', pady=(10, 2))
        e_email = ttk.Entry(self, textvariable=self.v_email, width=40, font=FONT)
        e_email.grid(row=2, column=0, sticky='we')
        tk.Label(self, text='Password', bg=CARD, font=('Segoe UI Semibold', 10)).grid(row=3, column=0, sticky='w', pady=(6, 2))
        e_pw = ttk.Entry(self, textvariable=self.v_pw, width=40, font=FONT, show='•')
        e_pw.grid(row=4, column=0, sticky='we')
        self.error = tk.Label(self, bg='#fef2f2', fg='#b91c1c', font=FONT, anchor='w', justify='left', wraplength=340)
        btns = tk.Frame(self, bg=CARD)
        btns.grid(row=6, column=0, sticky='e', pady=(14, 0))
        ttk.Button(btns, text='Cancel', command=self.destroy).pack(side='right')
        self.b_ok = ttk.Button(btns, text='Sign in', style='Accent.TButton', command=self.submit)
        self.b_ok.pack(side='right', padx=(0, 6))
        self.bind('<Return>', lambda e: self.submit())
        self.bind('<Escape>', lambda e: self.destroy())
        self.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - self.winfo_width()) // 2
        y = app.winfo_rooty() + (app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f'+{max(x, 0)}+{max(y, 0)}')
        self.grab_set()
        (e_pw if self.v_email.get() else e_email).focus_set()

    def submit(self):
        email, pw = self.v_email.get().strip(), self.v_pw.get()
        if not email or not pw:
            return self.show_error('Enter your email and password.')
        self.b_ok.state(['disabled'])
        self.config(cursor='watch')
        self.update_idletasks()
        try:
            self.app.cloud.sign_in(email, pw)
        except CloudError as ex:
            self.b_ok.state(['!disabled'])
            self.config(cursor='')
            return self.show_error(str(ex))
        self.ok = True
        self.destroy()

    def show_error(self, msg):
        self.error.config(text=msg, padx=8, pady=6)
        self.error.grid(row=5, column=0, sticky='we', pady=(8, 0))


class TextDialog(tk.Toplevel):
    """Simple read-only message window with scrollable details."""

    def __init__(self, app, title, summary, lines):
        super().__init__(app)
        self.title(title)
        self.configure(bg=CARD, padx=16, pady=12)
        self.transient(app)
        tk.Label(self, text=summary, bg=CARD, font=('Segoe UI Semibold', 11), anchor='w').pack(fill='x')
        if lines:
            frame = tk.Frame(self, bg=CARD)
            frame.pack(fill='both', expand=True, pady=(8, 0))
            txt = tk.Text(frame, width=80, height=min(len(lines), 14), font=FONT, wrap='word', relief='solid', bd=1)
            sb = ttk.Scrollbar(frame, command=txt.yview)
            txt.configure(yscrollcommand=sb.set)
            txt.pack(side='left', fill='both', expand=True)
            sb.pack(side='right', fill='y')
            txt.insert('1.0', '\n'.join(lines))
            txt.configure(state='disabled')
        ttk.Button(self, text='OK', command=self.destroy).pack(anchor='e', pady=(10, 0))
        self.bind('<Escape>', lambda e: self.destroy())
        self.bind('<Return>', lambda e: self.destroy())
        self.grab_set()


class App(tk.Tk):
    def __init__(self, db_path=DB_PATH):
        super().__init__()
        self.title('Item Checker')
        self.geometry('1180x720')
        self.minsize(380, 480)
        self.configure(bg=BG)
        self.cloud = Cloud(SUPABASE_URL, SUPABASE_ANON_KEY, SESSION_PATH) if SUPABASE_URL and SUPABASE_ANON_KEY else None
        self.store = Store(CACHE_PATH if self.cloud else db_path, seed=self.cloud is None)
        self.remote_version = self.polled_version = None
        self.closing = False
        self.sort_col, self.sort_dir = 'item_name', 'asc'
        self.last_result = None

        self.setup_style()
        self.build_menu()
        self.build_header()
        self.build_banner()
        self.build_body()
        self.build_statusbar()
        self.bind_keys()
        self.refresh()
        self.set_status(f'Database: {self.store.path}')
        self.protocol('WM_DELETE_WINDOW', self.on_close)
        if self.cloud:
            self.after(50, self.start_cloud)

    # ---------- shared mode ----------
    def start_cloud(self):
        self.set_status('Connecting to the online database…')
        self.update_idletasks()
        if not self.cloud.resume():
            dlg = LoginDialog(self)
            self.wait_window(dlg)
            if not dlg.ok:
                return self.on_close()
        try:
            self.pull()
        except CloudError as ex:
            messagebox.showerror('Item Checker', f'Could not load the online item list.\n\n{ex}\n\n'
                                 'Showing the copy saved on this PC. Changes will be sent once the connection is back.',
                                 parent=self)
        else:
            self.offer_upload()
        threading.Thread(target=self.poll_loop, daemon=True).start()
        self.after(1000, self.check_remote)

    def offer_upload(self):
        """First run in shared mode: offer to copy the items from this PC's items.db to the empty online list."""
        if self.store.totals()['total'] or not os.path.exists(DB_PATH):
            return
        try:
            src = sqlite3.connect(f'file:{DB_PATH}?mode=ro', uri=True)
            try:
                n = src.execute('SELECT COUNT(*) FROM items').fetchone()[0]
            finally:
                src.close()
        except sqlite3.Error:
            return
        if n and messagebox.askyesno('Upload your items?',
                                     f'The online item list is empty.\n\nUpload the {n:,} items from this PC '
                                     f'({os.path.basename(DB_PATH)}) so everyone can see them?', parent=self):
            self.store.restore_from(DB_PATH)
            self.saved()
            messagebox.showinfo('Upload finished', f'{self.store.totals()["total"]:,} items are now online.', parent=self)

    def pull(self):
        v = self.cloud.version()
        self.store.load_remote(self.cloud.fetch_items())
        self.remote_version = v
        self.refresh()
        self.set_status(f'Online · signed in as {self.cloud.email} · changes are shared')

    def saved(self):
        """Call after any change to the item list. In shared mode, sends the changes online first."""
        if self.cloud and self.cloud.session:
            self.config(cursor='watch')
            self.set_status('Saving…')
            self.update_idletasks()
            try:
                self.cloud.push(*self.store.changes())
            except CloudError as ex:
                messagebox.showerror('Could not save online',
                                     f'{ex}\n\nThe list will be reloaded with the latest online data.', parent=self)
            try:
                self.pull()
            except CloudError as ex:
                self.set_status(f'Offline: {ex}')
            finally:
                self.config(cursor='')
        self.refresh()

    def poll_loop(self):
        while not self.closing:
            time.sleep(4)
            try:
                self.polled_version = self.cloud.version()
            except Exception:
                self.polled_version = None

    def check_remote(self):
        """Reload when someone else changed the online list."""
        if self.closing:
            return
        v = self.polled_version
        if v is not None and v != self.remote_version and not self.grab_current():
            try:
                self.pull()
            except CloudError as ex:
                self.set_status(f'Offline: {ex}')
        self.after(1000, self.check_remote)

    def open_web(self):
        webbrowser.open(WEB_URL)

    def sign_out(self):
        if messagebox.askyesno('Sign out?', 'Sign out of the shared item list? The app will close.', parent=self):
            self.cloud.sign_out()
            self.on_close()

    # ---------- setup ----------
    def setup_style(self):
        s = ttk.Style(self)
        if 'vista' in s.theme_names():
            s.theme_use('vista')
        self.option_add('*Font', FONT)
        s.configure('.', font=FONT)
        s.configure('Treeview', rowheight=28, font=FONT)
        s.configure('Treeview.Heading', font=('Segoe UI Semibold', 10))
        s.configure('Accent.TButton', font=('Segoe UI Semibold', 10))
        s.configure('TNotebook.Tab', padding=(14, 5), font=('Segoe UI Semibold', 10))

    def build_menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label='Import CSV / Excel / .db backup…', accelerator='Ctrl+I', command=self.do_import)
        f.add_command(label='Download import template…', command=self.do_template)
        f.add_separator()
        f.add_command(label='Export to Excel (.xlsx)…', command=lambda: self.do_export('xlsx'))
        f.add_command(label='Export to CSV (.csv)…', command=lambda: self.do_export('csv'))
        f.add_command(label='SQLite backup (.db)…', command=lambda: self.do_export('db'))
        f.add_separator()
        if self.cloud:
            f.add_command(label='Open web version', command=self.open_web)
            f.add_command(label='Sign out…', command=self.sign_out)
        f.add_command(label='Open data folder', command=lambda: os.startfile(os.path.dirname(os.path.abspath(self.store.path))))
        f.add_separator()
        f.add_command(label='Exit', command=self.on_close)
        m.add_cascade(label='File', menu=f)
        it = tk.Menu(m, tearoff=0)
        it.add_command(label='Add Item', accelerator='Ctrl+N', command=self.add_item)
        it.add_command(label='View Item', accelerator='Enter', command=self.open_item)
        it.add_command(label='Edit Item', accelerator='Ctrl+E', command=self.edit_item)
        it.add_command(label='Toggle Active', accelerator='Space', command=self.toggle_item)
        it.add_command(label='Delete Item', accelerator='Del', command=self.delete_item)
        it.add_separator()
        it.add_command(label='Refresh', accelerator='F5', command=self.refresh)
        m.add_cascade(label='Items', menu=it)
        self.config(menu=m)
        self.row_menu = it

    def build_header(self):
        h = self.header = tk.Frame(self, bg=TEAL)
        h.pack(fill='x')
        left = tk.Frame(h, bg=TEAL)
        tk.Label(left, text='Item Checker', bg=TEAL, fg='white', font=('Segoe UI Semibold', 16)).pack(anchor='w')
        tk.Label(left, text='Stock count · SQLite', bg=TEAL, fg='#ccfbf1', font=('Segoe UI', 9)).pack(anchor='w')
        right = tk.Frame(h, bg=TEAL)
        ttk.Button(right, text='+ Add Item', style='Accent.TButton', command=self.add_item).pack(side='right', padx=(6, 0))
        exp = ttk.Menubutton(right, text='Export')
        em = tk.Menu(exp, tearoff=0)
        em.add_command(label='Excel (.xlsx)', command=lambda: self.do_export('xlsx'))
        em.add_command(label='CSV (.csv)', command=lambda: self.do_export('csv'))
        em.add_separator()
        em.add_command(label='SQLite backup (.db)', command=lambda: self.do_export('db'))
        exp['menu'] = em
        exp.pack(side='right', padx=(6, 0))
        ttk.Button(right, text='Import', command=self.do_import).pack(side='right')
        if self.cloud:
            ttk.Button(right, text='Open in Browser', command=self.open_web).pack(side='right', padx=(0, 6))

        # Title on the left and buttons on the right; on a narrow window the buttons move under the title.
        def relayout(e=None):
            lw, lh = left.winfo_reqwidth(), left.winfo_reqheight()
            rw, rh = right.winfo_reqwidth(), right.winfo_reqheight()
            left.place(x=16, y=10)
            if h.winfo_width() - 32 >= lw + rw + 16:
                right.place(relx=1, x=-16, y=10 + (lh - rh) // 2, anchor='ne')
                height = 20 + max(lh, rh)
            else:
                right.place(relx=0, x=16, y=18 + lh, anchor='nw')
                height = 28 + lh + rh
            if int(h.cget('height')) != height:
                h.configure(height=height)
        h.bind('<Configure>', relayout)
        h.after_idle(relayout)

    def build_banner(self):
        self.banner = tk.Frame(self, bg='#fffbeb', highlightbackground='#fcd34d', highlightthickness=1, padx=12, pady=8)
        tk.Label(self.banner, text='The app started with 5 sample items so you can try it out.', bg='#fffbeb', fg='#78350f').pack(side='left')
        ttk.Button(self.banner, text='Keep', command=self.keep_sample).pack(side='right')
        ttk.Button(self.banner, text='Clear sample data', command=self.clear_sample).pack(side='right', padx=(0, 6))

    def build_body(self):
        self.body = tk.Frame(self, bg=BG, padx=14, pady=10)
        self.body.pack(fill='both', expand=True)
        self.nb = ttk.Notebook(self.body)
        self.nb.pack(fill='both', expand=True)
        self.tab_search = tk.Frame(self.nb, bg=BG, padx=14, pady=10)
        self.tab_items = tk.Frame(self.nb, bg=BG, padx=10, pady=10)
        self.tab_sql = tk.Frame(self.nb, bg=BG, padx=10, pady=10)
        self.nb.add(self.tab_search, text='Item Search')
        self.nb.add(self.tab_items, text='All Items')
        # SQL Console hidden for now:
        # self.nb.add(self.tab_sql, text='SQL Console')
        self.build_search_tab()
        self.build_items_tab()
        self.build_sql_tab()

    def build_search_tab(self):
        t = self.tab_search

        # Search bar at top
        search_bar = tk.Frame(t, bg=BG)
        search_bar.pack(fill='x', pady=(4, 6))

        search_row = tk.Frame(search_bar, bg=BG)
        search_row.pack(fill='x')

        lbl_scan = tk.Label(search_row, text='Scan / Search Item:', bg=BG, fg='#0f172a',
                            font=('Segoe UI Semibold', 11))

        self.v_item_search = tk.StringVar()
        self.e_item_search = ttk.Entry(search_row, textvariable=self.v_item_search,
                                       font=('Segoe UI', 11), width=24)
        self.e_item_search.bind('<Return>', lambda e: self.do_item_search())
        self.e_item_search.bind('<KP_Enter>', lambda e: self.do_item_search())

        btn_box = tk.Frame(search_row, bg=BG)
        btn_search = ttk.Button(btn_box, text='Search', style='Accent.TButton',
                                command=self.do_item_search)
        btn_search.pack(side='left')

        btn_clear = ttk.Button(btn_box, text='Clear', command=self.clear_item_search)
        btn_clear.pack(side='left', padx=(4, 0))

        flow(search_row, [lbl_scan, self.e_item_search, btn_box], gap=8, align='center')

        self.l_search_msg = tk.Label(search_bar, text='', bg=BG, fg=MUTED, font=('Segoe UI', 9))
        self.l_search_msg.pack(pady=(4, 0))

        # Main content area
        self.search_content = tk.Frame(t, bg=BG)
        self.search_content.pack(fill='both', expand=True)

        # 1. Empty state
        self.search_empty = tk.Frame(self.search_content, bg=BG)
        self.search_empty.pack(fill='both', expand=True, pady=40)
        tk.Label(self.search_empty, text='No item selected', bg=BG, fg=MUTED,
                 font=('Segoe UI Semibold', 13)).pack()
        tk.Label(self.search_empty,
                 text='Scan a barcode or enter SKU / Item Name above to inspect and enter counts.',
                 bg=BG, fg=MUTED, font=('Segoe UI', 10)).pack(pady=(6, 0))

        # 2. Multiple matches selector
        self.search_matches = tk.Frame(self.search_content, bg=BG)
        tk.Label(self.search_matches, text='Multiple matches found — double-click an item to inspect:',
                 bg=BG, fg='#0f172a', font=('Segoe UI Semibold', 10)).pack(anchor='w', pady=(0, 6))
        match_table = tk.Frame(self.search_matches, bg=BG)
        match_table.pack(fill='both', expand=True)
        self.match_cols = [('item_name', 'Item Name', 0.45), ('sku', 'SKU', 0.25),
                           ('bin_location', 'Bin Location', 0.15), ('status', 'Count Result', 0.15)]
        self.match_tree = ttk.Treeview(match_table, columns=[c[0] for c in self.match_cols],
                                       show='headings', selectmode='browse', height=8)
        for key, title, _ in self.match_cols:
            self.match_tree.heading(key, text=title, anchor='w')
            self.match_tree.column(key, width=120, anchor='w')
        sb = ttk.Scrollbar(match_table, command=self.match_tree.yview)
        self.match_tree.configure(yscrollcommand=sb.set)
        self.match_tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        self.match_tree.bind('<Double-1>', lambda e: self.on_match_selected())
        self.match_tree.bind('<Return>', lambda e: self.on_match_selected())

        # 3. Card view
        self.search_card_wrap = tk.Frame(self.search_content, bg=BG)
        self.card_sub = tk.Frame(self.search_card_wrap, bg=BG)
        self.card_sub.pack(anchor='n', pady=6)
        self.card_sub.columnconfigure(0, weight=1)
        self.search_card = ItemCardFrame(self.card_sub, app=self, show_close=False,
                                         on_deleted=self.on_search_card_deleted)
        self.search_card.grid(row=0, column=0, sticky='nsew')

        def on_card_wrap_resize(e):
            avail = max(280, e.width - 24)
            w = max(280, min(avail, 460))
            self.card_sub.columnconfigure(0, minsize=w)
        self.search_card_wrap.bind('<Configure>', on_card_wrap_resize)

    def do_item_search(self):
        query_text = self.v_item_search.get().strip()
        if not query_text:
            self.clear_item_search()
            return

        # 1. Exact match on SKU first
        exact = self.store.find_sku(query_text)
        if exact:
            self.show_search_item(exact['id'])
            self.l_search_msg.config(text=f'Showing exact SKU match: {exact["sku"]}', fg=TEAL)
            return

        # 2. General search across fields
        like = '%' + re.sub(r'([\\%_])', r'\\\1', query_text) + '%'
        rows = self.store.conn.execute(
            "SELECT id, item_name, sku, item_ref, warehouse, bin_location, status "
            "FROM v_item_check WHERE (sku LIKE ? ESCAPE '\\' OR item_name LIKE ? ESCAPE '\\' "
            "OR item_ref LIKE ? ESCAPE '\\' OR warehouse LIKE ? ESCAPE '\\' OR bin_location LIKE ? ESCAPE '\\') "
            "ORDER BY item_name COLLATE NOCASE LIMIT 50",
            (like, like, like, like, like)
        ).fetchall()

        if len(rows) == 1:
            self.show_search_item(rows[0]['id'])
            self.l_search_msg.config(text=f'Found: {rows[0]["sku"]} ({rows[0]["item_name"]})', fg=TEAL)
        elif len(rows) > 1:
            self.show_search_matches(rows)
            self.l_search_msg.config(text=f'Found {len(rows)} matching items. Select one below.', fg='#0f172a')
        else:
            self.show_search_empty()
            self.l_search_msg.config(text=f'No match found for "{query_text}"', fg=RED)

    def show_search_item(self, item_id):
        self.search_empty.pack_forget()
        self.search_matches.pack_forget()
        self.search_card_wrap.pack_forget()
        if self.search_card.load_item(item_id):
            self.search_card_wrap.pack(fill='both', expand=True)

    def show_search_matches(self, rows):
        self.search_empty.pack_forget()
        self.search_card_wrap.pack_forget()
        self.search_matches.pack_forget()
        self.match_tree.delete(*self.match_tree.get_children())
        for r in rows:
            self.match_tree.insert('', 'end', iid=str(r['id']),
                                   values=(r['item_name'], r['sku'], r['bin_location'] or '—', r['status']))
        if rows:
            first_id = str(rows[0]['id'])
            self.match_tree.selection_set(first_id)
            self.match_tree.focus(first_id)
        self.search_matches.pack(fill='both', expand=True, pady=10)

    def show_search_empty(self):
        self.search_card_wrap.pack_forget()
        self.search_matches.pack_forget()
        self.search_empty.pack_forget()
        self.search_empty.pack(fill='both', expand=True, pady=40)

    def on_match_selected(self):
        sel = self.match_tree.selection()
        if sel:
            item_id = int(sel[0])
            self.show_search_item(item_id)

    def clear_item_search(self):
        self.v_item_search.set('')
        self.l_search_msg.config(text='')
        self.show_search_empty()
        self.e_item_search.focus_set()

    def on_search_card_deleted(self, deleted_id):
        self.clear_item_search()
        self.refresh()

    def show_in_search_tab(self, item_id):
        r = self.store.get(item_id)
        if r:
            self.v_item_search.set(r['sku'])
            self.nb.select(self.tab_search)
            self.show_search_item(item_id)
            self.l_search_msg.config(text=f'Showing item: {r["sku"]}', fg=TEAL)

    def stat_card(self, parent, title, color=None, on_click=None):
        """Builds a card (placed later by the reflow in build_items_tab) and returns its value label."""
        card = tk.Frame(parent, bg=CARD, highlightbackground='#e2e8f0', highlightthickness=1, padx=14, pady=10)
        t = tk.Label(card, text=title.upper(), bg=CARD, fg=color or MUTED, font=('Segoe UI Semibold', 8))
        t.pack(anchor='w')
        v = tk.Label(card, text='0', bg=CARD, fg=color or '#0f172a', font=('Segoe UI Semibold', 18))
        v.pack(anchor='w')
        if on_click:
            for w in (card, t, v):
                w.bind('<Button-1>', lambda e: on_click())
                w.configure(cursor='hand2')
        return v

    def build_items_tab(self):
        t = self.tab_items
        # Stat cards: 5 across on a wide window, 3 or 2 across on a narrow one.
        cards = tk.Frame(t, bg=BG)
        cards.pack(fill='x')
        self.s_total = self.stat_card(cards, 'Total Items', on_click=lambda: self.quick_filter('all', 'all'))
        self.s_active = self.stat_card(cards, 'Active', on_click=lambda: self.quick_filter('1', 'all'))
        self.s_phys = self.stat_card(cards, 'Physical QTY')
        self.s_tocount = self.stat_card(cards, 'To count', AMBER, on_click=lambda: self.quick_filter('all', 'TO COUNT'))
        self.s_mismatch = self.stat_card(cards, 'Mismatch', RED, on_click=lambda: self.quick_filter('all', 'MISMATCH'))
        card_frames = [v.master for v in (self.s_total, self.s_active, self.s_phys, self.s_tocount, self.s_mismatch)]
        self._card_cols = None

        def reflow_cards(e):
            n = 5 if e.width >= 760 else 3 if e.width >= 480 else 2
            if n == self._card_cols:
                return
            self._card_cols = n
            for i in range(5):
                cards.columnconfigure(i, weight=1 if i < n else 0, uniform='c' if i < n else '')
            for i, c in enumerate(card_frames):
                row, col = divmod(i, n)
                c.grid(row=row, column=col, sticky='nsew', padx=(0 if col == 0 else 8, 0), pady=(0 if row == 0 else 8, 0))
        cards.bind('<Configure>', reflow_cards)

        bar = tk.Frame(t, bg=BG)
        bar.pack(fill='x', pady=10)
        search = tk.Frame(bar, bg=BG)
        tk.Label(search, text='Search', bg=BG, fg=MUTED).pack(side='left')
        self.v_search = tk.StringVar()
        self.e_search = ttk.Entry(search, textvariable=self.v_search, width=30)
        self.e_search.pack(side='left', padx=(6, 0))
        self.v_search.trace_add('write', lambda *_: self.refresh())
        self.active_opts = {'All items': 'all', 'Active only': '1', 'Inactive only': '0'}
        self.v_active = tk.StringVar(value='All items')
        cb = ttk.Combobox(bar, textvariable=self.v_active, values=list(self.active_opts), state='readonly', width=14)
        cb.bind('<<ComboboxSelected>>', lambda e: self.refresh())
        self.v_statusf = tk.StringVar(value='All statuses')
        cb2 = ttk.Combobox(bar, textvariable=self.v_statusf, values=['All statuses'] + STATUSES, state='readonly', width=18)
        cb2.bind('<<ComboboxSelected>>', lambda e: self.refresh())
        reset = ttk.Button(bar, text='Reset', command=lambda: self.quick_filter('all', 'all'))
        self.l_count = tk.Label(bar, bg=BG, fg=MUTED)
        flow(bar, [search, cb, cb2, reset, self.l_count], gap=8)

        # The list shows Item Name, SKU and ItemRef only; opening an item shows its card with every field.
        table = tk.Frame(t, bg=BG)
        table.pack(fill='both', expand=True)
        self.cols = [('item_name', 'Item Name', 0.5), ('sku', 'SKU', 0.25), ('item_ref', 'ItemRef', 0.25)]
        self.tree = ttk.Treeview(table, columns=[c[0] for c in self.cols], show='headings', selectmode='browse')
        for key, title, _ in self.cols:
            self.tree.heading(key, text=title, anchor='w', command=lambda k=key: self.sort_by(k))
            self.tree.column(key, width=100, minwidth=60, anchor='w', stretch=True)
        self.tree.tag_configure('inactive', foreground='#94a3b8')
        sb = ttk.Scrollbar(table, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        def fit_columns(e):
            for key, _, share in self.cols:
                self.tree.column(key, width=max(int(e.width * share), 60))
        self.tree.bind('<Configure>', fit_columns)
        self.tree.bind('<Double-1>', lambda e: self.open_item() if self.tree.identify_row(e.y) else None)
        self.tree.bind('<Return>', lambda e: self.open_item())
        self.tree.bind('<space>', lambda e: self.toggle_item())
        self.tree.bind('<Delete>', lambda e: self.delete_item())
        self.tree.bind('<Button-3>', self.on_right_click)
        self.tree.bind('<<TreeviewSelect>>', lambda e: self.update_buttons())

        self.empty = tk.Label(self.tree, bg=CARD, fg=MUTED)

        actions = tk.Frame(t, bg=BG)
        actions.pack(side='bottom', fill='x', pady=(8, 0), before=table)  # keeps the buttons visible on a short window
        self.b_open = ttk.Button(actions, text='View', style='Accent.TButton', command=self.open_item)
        self.b_edit = ttk.Button(actions, text='Edit', command=self.edit_item)
        self.b_toggle = ttk.Button(actions, text='Toggle Active', command=self.toggle_item)
        self.b_delete = ttk.Button(actions, text='Delete', command=self.delete_item)
        tip = tk.Label(actions, text='Tip: double-click an item to view it · Space toggles Active · Del deletes',
                       bg=BG, fg=MUTED, font=('Segoe UI', 9))
        flow(actions, [self.b_open, self.b_edit, self.b_toggle, self.b_delete, tip])

    def build_sql_tab(self):
        t = self.tab_sql
        presets = tk.Frame(t, bg=BG)
        presets.pack(fill='x')
        for label, sql in PRESETS:
            ttk.Button(presets, text=label, command=lambda s=sql: self.run_preset(s)).pack(side='left', padx=(0, 6))
        box = tk.Frame(t, bg=BG, pady=8)
        box.pack(fill='x')
        self.sql_text = tk.Text(box, height=8, font=('Consolas', 11), wrap='none', undo=True,
                                relief='solid', bd=1, padx=8, pady=6)
        self.sql_text.pack(fill='x')
        self.sql_text.insert('1.0', "SELECT * FROM v_item_check WHERE status = 'MISMATCH';")
        self.sql_text.bind('<Control-Return>', lambda e: (self.run_sql(), 'break')[1])
        row = tk.Frame(t, bg=BG)
        row.pack(fill='x')
        ttk.Button(row, text='Run  (Ctrl+Enter)', style='Accent.TButton', command=self.run_sql).pack(side='left')
        self.b_sql_csv = ttk.Button(row, text='Save result as CSV…', command=self.save_sql_csv, state='disabled')
        self.b_sql_csv.pack(side='left', padx=(6, 0))
        self.l_sql = tk.Label(row, bg=BG, fg=MUTED)
        self.l_sql.pack(side='right')
        tk.Label(t, text='Table: items   ·   View: v_item_check (adds status: MATCH, MISMATCH, TO COUNT or NO STOCK)   ·   Changes made here are saved to the database.',
                 bg=BG, fg=MUTED, font=('Segoe UI', 9)).pack(anchor='w', pady=(4, 6))
        self.sql_err = tk.Label(t, bg='#fef2f2', fg='#b91c1c', font=('Consolas', 10), anchor='w', justify='left', padx=8, pady=6)
        res = tk.Frame(t, bg=BG)
        res.pack(fill='both', expand=True)
        self.sql_tree = ttk.Treeview(res, show='headings')
        ys = ttk.Scrollbar(res, command=self.sql_tree.yview)
        xs = ttk.Scrollbar(res, orient='horizontal', command=self.sql_tree.xview)
        self.sql_tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.sql_tree.grid(row=0, column=0, sticky='nsew')
        ys.grid(row=0, column=1, sticky='ns')
        xs.grid(row=1, column=0, sticky='ew')
        res.rowconfigure(0, weight=1)
        res.columnconfigure(0, weight=1)

    def build_statusbar(self):
        self.statusbar = tk.Label(self, bg='#e2e8f0', fg='#334155', anchor='w', padx=12, pady=3, font=('Segoe UI', 9))
        self.statusbar.pack(fill='x', side='bottom')

    def bind_keys(self):
        self.bind('<Control-n>', lambda e: self.add_item())
        self.bind('<Control-e>', lambda e: self.edit_item())
        self.bind('<Control-i>', lambda e: self.do_import())
        self.bind('<Control-f>', lambda e: (self.nb.select(self.tab_search), self.e_item_search.focus_set()))
        self.bind('<F5>', lambda e: self.refresh())

    # ---------- helpers ----------
    def set_status(self, msg):
        self.statusbar.config(text=f'  {msg}')

    def selected_id(self):
        sel = self.tree.selection()
        return int(sel[0]) if sel else None

    def update_buttons(self):
        state = 'normal' if self.selected_id() else 'disabled'
        for b in (self.b_open, self.b_edit, self.b_toggle, self.b_delete):
            b.configure(state=state)

    def quick_filter(self, active, status):
        self.v_active.set({v: k for k, v in self.active_opts.items()}[active])
        self.v_statusf.set('All statuses' if status == 'all' else status)
        self.nb.select(self.tab_items)
        self.v_search.set('')  # triggers refresh

    def sort_by(self, col):
        if self.sort_col == col:
            self.sort_dir = 'desc' if self.sort_dir == 'asc' else 'asc'
        else:
            self.sort_col, self.sort_dir = col, 'asc'
        self.refresh()

    # ---------- render ----------
    def refresh(self):
        s = self.store
        t = s.totals()
        self.s_total.config(text=f'{t["total"]:,}')
        self.s_active.config(text=f'{t["active"]:,}')
        self.s_phys.config(text=f'{t["physical"]:,}')
        self.s_tocount.config(text=f'{t["to_count"]:,}')
        self.s_mismatch.config(text=f'{t["mismatch"]:,}')

        status = self.v_statusf.get()
        rows = s.list_items(self.v_search.get().strip(), self.active_opts.get(self.v_active.get(), 'all'),
                            status if status in STATUSES else 'all', self.sort_col, self.sort_dir)
        keep = self.selected_id()
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            self.tree.insert('', 'end', iid=str(r['id']), tags=() if r['active'] else ('inactive',),
                             values=(r['item_name'], r['sku'], r['item_ref'] or '—'))
        if keep and self.tree.exists(str(keep)):
            self.tree.selection_set(str(keep))
            self.tree.see(str(keep))
        self.l_count.config(text=f'Showing {len(rows):,} of {t["total"]:,}')

        for key, title, _ in self.cols:
            arrow = (' ▲' if self.sort_dir == 'asc' else ' ▼') if key == self.sort_col else ''
            self.tree.heading(key, text=title + arrow)

        if rows:
            self.empty.place_forget()
        else:
            self.empty.config(text='No items match the current filters.' if t['total']
                              else 'No items yet. Click "+ Add Item" or Import a file.')
            self.empty.place(relx=0.5, rely=0.4, anchor='center')

        if s.sample_flag:
            self.banner.pack(fill='x', padx=14, pady=(10, 0), after=self.header)
        else:
            self.banner.pack_forget()
        self.update_buttons()
        if hasattr(self, 'search_card') and self.search_card.item_id:
            self.search_card.fill()

    # ---------- item actions ----------
    def add_item(self):
        ItemDialog(self)

    def open_item(self):
        item_id = self.selected_id()
        if item_id:
            self.show_in_search_tab(item_id)

    def edit_item(self):
        item_id = self.selected_id()
        if item_id:
            item = self.store.get(item_id)
            if item:
                ItemDialog(self, item)

    def toggle_item(self):
        item_id = self.selected_id()
        if item_id:
            self.store.toggle_active(item_id)
            self.saved()

    def delete_item(self):
        item_id = self.selected_id()
        if not item_id:
            return
        r = self.store.get(item_id)
        if r and messagebox.askyesno('Delete item?', f'{r["item_name"]} ({r["sku"]}) will be permanently removed.',
                                     icon='warning', parent=self):
            self.store.delete(item_id)
            self.saved()
            self.set_status(f'Deleted {r["sku"]}')

    def on_right_click(self, e):
        row = self.tree.identify_row(e.y)
        if row:
            self.tree.selection_set(row)
            self.row_menu.tk_popup(e.x_root, e.y_root)

    def clear_sample(self):
        if messagebox.askyesno('Clear sample data?', 'All items currently in the list will be removed so you can start fresh.',
                               icon='warning', parent=self):
            self.store.clear_all()
            self.saved()
            self.set_status('Sample data cleared')

    def keep_sample(self):
        self.store.sample_flag = False
        self.refresh()

    # ---------- import / export ----------
    def do_import(self):
        path = filedialog.askopenfilename(parent=self, title='Import items', filetypes=[
            ('Item files', '*.xlsx *.xlsm *.csv *.db *.sqlite *.sqlite3'), ('Excel', '*.xlsx *.xlsm'), ('CSV', '*.csv'),
            ('SQLite backup', '*.db *.sqlite *.sqlite3'), ('All files', '*.*')])
        if not path:
            return
        try:
            if os.path.splitext(path)[1].lower() in ('.db', '.sqlite', '.sqlite3'):
                if not messagebox.askyesno('Restore backup?', 'Restoring a .db backup REPLACES ALL current items.\n\nContinue?',
                                           icon='warning', parent=self):
                    return
                self.store.restore_from(path)
                self.saved()
                n = self.store.totals()['total']
                self.set_status(f'Backup restored from {os.path.basename(path)}')
                messagebox.showinfo('Backup restored', f'{n:,} items loaded.', parent=self)
                return
            added, updated, skipped = self.store.import_rows(read_table_file(path))
        except Exception as ex:
            messagebox.showerror('Import failed', str(ex), parent=self)
            return
        self.saved()
        summary = f'{added:,} added, {updated:,} updated, {len(skipped):,} skipped.'
        self.set_status(f'Imported {os.path.basename(path)}: {summary}')
        lines = [f'Row {n}{f" ({sku})" if sku else ""}: {why}' for n, sku, why in skipped]
        TextDialog(self, 'Import finished', summary, lines)

    def ask_save(self, title, ext, label):
        return filedialog.asksaveasfilename(parent=self, title=title, defaultextension=ext,
                                            initialfile=f'ItemChecker_{today()}{ext}', filetypes=[(label, f'*{ext}')])

    def do_export(self, kind):
        try:
            if kind == 'db':
                path = self.ask_save('Save SQLite backup', '.db', 'SQLite database')
                if path:
                    self.store.backup_to(path)
            else:
                if kind == 'xlsx' and openpyxl is None:
                    raise RuntimeError('Excel export needs openpyxl.  Run:  pip install openpyxl')
                path = self.ask_save('Export items', f'.{kind}', 'Excel workbook' if kind == 'xlsx' else 'CSV file')
                if path:
                    rows = self.store.export_rows()
                    if kind == 'xlsx':
                        write_xlsx(path, rows)
                    else:
                        write_csv(path, EXPORT_HEADERS,
                                  [export_values(r, 'TRUE' if r['active'] else 'FALSE') for r in rows])
        except Exception as ex:
            messagebox.showerror('Export failed', str(ex), parent=self)
            return
        if path:
            self.set_status(f'Saved {path}')

    def do_template(self):
        if openpyxl is None:
            messagebox.showerror('Template', 'Excel support needs openpyxl.  Run:  pip install openpyxl', parent=self)
            return
        path = filedialog.asksaveasfilename(parent=self, title='Save import template', defaultextension='.xlsx',
                                            initialfile='ItemChecker_Import_Template.xlsx', filetypes=[('Excel workbook', '*.xlsx')])
        if path:
            try:
                write_template(path)
                self.set_status(f'Template saved: {path}')
            except Exception as ex:
                messagebox.showerror('Template', str(ex), parent=self)

    # ---------- SQL console ----------
    def run_preset(self, sql):
        self.sql_text.delete('1.0', 'end')
        self.sql_text.insert('1.0', sql)
        self.run_sql()

    def run_sql(self):
        sql = self.sql_text.get('1.0', 'end').strip()
        if not sql:
            return
        self.sql_err.pack_forget()
        started = datetime.datetime.now()
        try:
            cols, rows, changed, any_change = self.store.run_sql(sql)
        except sqlite3.Error as ex:
            if self.store.conn.in_transaction:
                self.store.conn.execute('ROLLBACK')
            self.sql_err.config(text=f'Error: {ex}')
            self.sql_err.pack(fill='x', before=self.sql_tree.master, pady=(0, 6))
            self.show_sql_result(None, None)
            self.l_sql.config(text='')
            return
        ms = (datetime.datetime.now() - started).total_seconds() * 1000
        self.show_sql_result(cols, rows)
        parts = []
        if cols is not None:
            parts.append(f'{len(rows):,} row(s)')
        if changed:
            parts.append(f'{changed:,} changed')
        elif cols is None:
            parts.append('Done')
        parts.append(f'{ms:.1f} ms')
        self.l_sql.config(text='  ·  '.join(parts))
        if any_change:
            self.saved()

    def show_sql_result(self, cols, rows):
        tv = self.sql_tree
        tv.delete(*tv.get_children())
        self.last_result = (cols, rows) if cols is not None else None
        self.b_sql_csv.configure(state='normal' if self.last_result else 'disabled')
        if cols is None:
            tv['columns'] = ()
            return
        tv['columns'] = [f'c{i}' for i in range(len(cols))]
        for i, c in enumerate(cols):
            numeric = any(isinstance(r[i], (int, float)) for r in rows[:50]) and all(
                r[i] is None or isinstance(r[i], (int, float)) for r in rows[:50])
            tv.heading(f'c{i}', text=c, anchor='e' if numeric else 'w')
            tv.column(f'c{i}', anchor='e' if numeric else 'w', width=max(90, min(360, len(c) * 10 + 40)), stretch=True)
        for r in rows[:5000]:
            tv.insert('', 'end', values=['NULL' if v is None else (f'{v:,}' if isinstance(v, int) else v) for v in r])

    def save_sql_csv(self):
        if not self.last_result:
            return
        path = filedialog.asksaveasfilename(parent=self, title='Save query result', defaultextension='.csv',
                                            initialfile=f'ItemChecker_query_{today()}.csv', filetypes=[('CSV file', '*.csv')])
        if path:
            write_csv(path, *self.last_result)
            self.set_status(f'Saved {path}')

    def on_close(self):
        self.closing = True
        self.store.close()
        self.destroy()


def main():
    if sys.platform == 'win32':
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    try:
        app = App()
    except sqlite3.Error as ex:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror('Item Checker', f'Could not open the database:\n{DB_PATH}\n\n{ex}')
        return
    app.mainloop()


if __name__ == '__main__':
    main()
