-- Item Checker · SQLite schema
-- One table, one trigger, one view.

CREATE TABLE IF NOT EXISTS items (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  item_name       TEXT    NOT NULL,
  sku             TEXT    NOT NULL UNIQUE COLLATE NOCASE,
  physical_stock  INTEGER NOT NULL DEFAULT 0 CHECK (physical_stock  >= 0),
  committed_stock INTEGER NOT NULL DEFAULT 0 CHECK (committed_stock >= 0),
  active          INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
  created_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
  updated_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE TRIGGER IF NOT EXISTS trg_items_updated AFTER UPDATE ON items
BEGIN
  UPDATE items SET updated_at = datetime('now','localtime') WHERE id = OLD.id;
END;

-- Status rules, checked in order:
--   OVER-COMMITTED  committed > physical
--   NO STOCK        physical = 0
--   LOW             available <= 5
--   OK              everything else
CREATE VIEW IF NOT EXISTS v_item_check AS
SELECT id, item_name, sku, physical_stock, committed_stock,
       (physical_stock - committed_stock) AS available_stock, active,
       CASE WHEN committed_stock > physical_stock       THEN 'OVER-COMMITTED'
            WHEN physical_stock = 0                     THEN 'NO STOCK'
            WHEN physical_stock - committed_stock <= 5  THEN 'LOW'
            ELSE 'OK' END AS status,
       updated_at
FROM items;
