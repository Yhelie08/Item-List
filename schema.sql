-- Item Checker · SQLite schema
-- One table, one trigger, one view.

CREATE TABLE IF NOT EXISTS items (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  item_name       TEXT    NOT NULL,
  sku             TEXT    NOT NULL UNIQUE COLLATE NOCASE,
  item_ref        TEXT    NOT NULL DEFAULT '',
  warehouse       TEXT    NOT NULL DEFAULT '',
  bin_location    TEXT    NOT NULL DEFAULT '',
  physical_stock  INTEGER NOT NULL DEFAULT 0 CHECK (physical_stock  >= 0),
  verifier1_count INTEGER CHECK (verifier1_count >= 0),  -- NULL = not counted yet
  verifier2_count INTEGER CHECK (verifier2_count >= 0),
  active          INTEGER NOT NULL DEFAULT 1 CHECK (active IN (0,1)),
  created_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime')),
  updated_at      TEXT    NOT NULL DEFAULT (datetime('now','localtime'))
);

CREATE TRIGGER IF NOT EXISTS trg_items_updated AFTER UPDATE ON items
BEGIN
  UPDATE items SET updated_at = datetime('now','localtime') WHERE id = OLD.id;
END;

-- Status rules, checked in order:
--   NO STOCK   physical = 0 (the count is typed straight into Physical Stock)
--   TO COUNT   verifier 1 or verifier 2 has not counted yet
--   MATCH      both verifier counts equal physical
--   MISMATCH   everything else
CREATE VIEW IF NOT EXISTS v_item_check AS
SELECT id, item_name, sku, item_ref, warehouse, bin_location, physical_stock,
       verifier1_count, verifier2_count, active,
       CASE WHEN physical_stock = 0 THEN 'NO STOCK'
            WHEN verifier1_count IS NULL OR verifier2_count IS NULL THEN 'TO COUNT'
            WHEN verifier1_count = physical_stock AND verifier2_count = physical_stock THEN 'MATCH'
            ELSE 'MISMATCH' END AS status,
       updated_at
FROM items;
