"""
db.py
-----
SQLite data layer for the DMD Billing Management System.

Design notes
============
- Billing records are uniquely identified by (site_code, cycle_month).
  This is enforced with a UNIQUE constraint so re-importing the same
  Excel can never create a duplicate row.
- ISDP Status / IBuy Status / Remarks are the fields the user edits
  daily inside the app. Excel re-imports never touch these fields
  unless the user explicitly ticks "overwrite manual fields" on the
  Import page.
- Master data is keyed by `column1` (e.g. "DC-Document_Controller-001-M3"),
  which already encodes Uniq-ID + Period and is unique per source row group.
"""

import sqlite3
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "dmd_billing.db"

# Fields that represent "daily manual updates". Excel imports must never
# silently overwrite these unless the user opts in explicitly.
MANUAL_BILLING_FIELDS = ["isdp_status", "ibuy_status", "remarks"]

# All billing columns that CAN come from an Excel import (excludes id/timestamps).
BILLING_IMPORT_FIELDS = [
    "site_code", "cycle_month", "po", "name", "project_assign", "site_name",
    "regional", "sub_contract_no", "pr_no", "po_line_no", "shipment_no",
    "item_code", "item_description", "unit_price", "requested_qty",
    "payment_coef", "total_price", "acceptance_date", "status",
    "isdp_status", "remarks", "ibuy_status", "id_project", "so_number",
    "invoice_number", "invoice_date", "invoice_month",
]

MASTER_IMPORT_FIELDS = [
    "uniq_id", "column1", "name_id", "account", "region", "period",
    "period_date", "period_status", "productivity_range", "productivity",
    "payment_coef",
]


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@contextmanager
def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with get_conn() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS billing (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            site_code TEXT NOT NULL,
            cycle_month TEXT NOT NULL,
            po TEXT,
            name TEXT,
            project_assign TEXT,
            site_name TEXT,
            regional TEXT,
            sub_contract_no TEXT,
            pr_no TEXT,
            po_line_no TEXT,
            shipment_no TEXT,
            item_code TEXT,
            item_description TEXT,
            unit_price REAL,
            requested_qty REAL,
            payment_coef REAL,
            total_price REAL,
            acceptance_date TEXT,
            status TEXT,
            isdp_status TEXT,
            remarks TEXT,
            ibuy_status TEXT,
            id_project TEXT,
            so_number TEXT,
            invoice_number TEXT,
            invoice_date TEXT,
            invoice_month TEXT,
            created_at TEXT,
            updated_at TEXT,
            last_imported_at TEXT,
            last_manual_edit_at TEXT,
            UNIQUE(site_code, cycle_month)
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS master_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            uniq_id TEXT NOT NULL,
            column1 TEXT NOT NULL UNIQUE,
            name_id TEXT,
            account TEXT,
            region TEXT,
            period TEXT,
            period_date TEXT,
            period_status TEXT,
            productivity_range TEXT,
            productivity TEXT,
            payment_coef REAL,
            created_at TEXT,
            updated_at TEXT,
            last_imported_at TEXT
        )
        """)
        conn.execute("""
        CREATE TABLE IF NOT EXISTS import_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT,
            kind TEXT,
            filename TEXT,
            sheet TEXT,
            inserted INTEGER,
            updated INTEGER,
            skipped_duplicates INTEGER,
            manual_fields_overwritten INTEGER,
            detail TEXT
        )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_billing_site ON billing(site_code)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_billing_cycle ON billing(cycle_month)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_master_uniqid ON master_data(uniq_id)")


# ---------------------------------------------------------------------------
# Billing: upsert (used by importer)
# ---------------------------------------------------------------------------

def upsert_billing_row(conn, row: dict, overwrite_manual: bool) -> str:
    """
    Insert or update a single billing row keyed on (site_code, cycle_month).
    Returns 'inserted' or 'updated'.
    Only non-null incoming values overwrite existing non-manual fields.
    Manual fields (isdp_status, ibuy_status, remarks) are only overwritten
    when overwrite_manual=True.
    """
    site_code = row.get("site_code")
    cycle_month = row.get("cycle_month")
    existing = conn.execute(
        "SELECT * FROM billing WHERE site_code = ? AND cycle_month = ?",
        (site_code, cycle_month),
    ).fetchone()

    ts = now()

    if existing is None:
        cols = BILLING_IMPORT_FIELDS
        values = [row.get(c) for c in cols]
        placeholders = ", ".join(["?"] * len(cols))
        conn.execute(
            f"""INSERT INTO billing ({", ".join(cols)}, created_at, updated_at, last_imported_at)
                VALUES ({placeholders}, ?, ?, ?)""",
            values + [ts, ts, ts],
        )
        return "inserted"
    else:
        set_clauses = []
        values = []
        for c in BILLING_IMPORT_FIELDS:
            if c in ("site_code", "cycle_month"):
                continue
            if c in MANUAL_BILLING_FIELDS and not overwrite_manual:
                continue
            new_val = row.get(c)
            if new_val is None or (isinstance(new_val, str) and new_val.strip() == ""):
                continue  # never blank out existing data with an empty import value
            set_clauses.append(f"{c} = ?")
            values.append(new_val)
        set_clauses.append("updated_at = ?")
        values.append(ts)
        set_clauses.append("last_imported_at = ?")
        values.append(ts)
        values.append(existing["id"])
        conn.execute(
            f"UPDATE billing SET {', '.join(set_clauses)} WHERE id = ?", values
        )
        return "updated"


def import_billing_rows(rows: list, overwrite_manual: bool = False) -> dict:
    inserted = updated = 0
    with get_conn() as conn:
        for row in rows:
            result = upsert_billing_row(conn, row, overwrite_manual)
            if result == "inserted":
                inserted += 1
            else:
                updated += 1
    return {"inserted": inserted, "updated": updated}


# ---------------------------------------------------------------------------
# Master data: upsert
# ---------------------------------------------------------------------------

def upsert_master_row(conn, row: dict) -> str:
    column1 = row.get("column1")
    existing = conn.execute(
        "SELECT id FROM master_data WHERE column1 = ?", (column1,)
    ).fetchone()
    ts = now()
    if existing is None:
        cols = MASTER_IMPORT_FIELDS
        values = [row.get(c) for c in cols]
        placeholders = ", ".join(["?"] * len(cols))
        conn.execute(
            f"""INSERT INTO master_data ({", ".join(cols)}, created_at, updated_at, last_imported_at)
                VALUES ({placeholders}, ?, ?, ?)""",
            values + [ts, ts, ts],
        )
        return "inserted"
    else:
        set_clauses = []
        values = []
        for c in MASTER_IMPORT_FIELDS:
            if c == "column1":
                continue
            new_val = row.get(c)
            if new_val is None or (isinstance(new_val, str) and new_val.strip() == ""):
                continue
            set_clauses.append(f"{c} = ?")
            values.append(new_val)
        set_clauses.append("updated_at = ?")
        values.append(ts)
        set_clauses.append("last_imported_at = ?")
        values.append(ts)
        values.append(existing["id"])
        conn.execute(
            f"UPDATE master_data SET {', '.join(set_clauses)} WHERE id = ?", values
        )
        return "updated"


def import_master_rows(rows: list) -> dict:
    inserted = updated = 0
    with get_conn() as conn:
        for row in rows:
            result = upsert_master_row(conn, row)
            if result == "inserted":
                inserted += 1
            else:
                updated += 1
    return {"inserted": inserted, "updated": updated}


def log_import(kind, filename, sheet, inserted, updated, skipped_duplicates, manual_overwritten, detail=""):
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO import_log
               (ts, kind, filename, sheet, inserted, updated, skipped_duplicates,
                manual_fields_overwritten, detail)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (now(), kind, filename, sheet, inserted, updated, skipped_duplicates,
             int(manual_overwritten), detail),
        )


def get_import_log(limit=50):
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM import_log ORDER BY id DESC LIMIT ?", (limit,)
        ).fetchall()
        return [dict(r) for r in rows]


# ---------------------------------------------------------------------------
# Billing: read / update (manual daily edits)
# ---------------------------------------------------------------------------

def get_all_billing():
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM billing ORDER BY site_code, cycle_month").fetchall()
        return [dict(r) for r in rows]


def get_billing_by_id(record_id: int):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM billing WHERE id = ?", (record_id,)).fetchone()
        return dict(row) if row else None


def update_billing_manual(record_id: int, isdp_status: str, ibuy_status: str, remarks: str):
    """The main 'daily update' entry point used by the Edit Billing form."""
    ts = now()
    with get_conn() as conn:
        conn.execute(
            """UPDATE billing
               SET isdp_status = ?, ibuy_status = ?, remarks = ?,
                   updated_at = ?, last_manual_edit_at = ?
               WHERE id = ?""",
            (isdp_status, ibuy_status, remarks, ts, ts, record_id),
        )


def update_billing_full(record_id: int, fields: dict):
    """Update any subset of editable billing fields from the Edit form."""
    ts = now()
    allowed = set(BILLING_IMPORT_FIELDS)
    set_clauses = []
    values = []
    manual_touch = False
    for k, v in fields.items():
        if k not in allowed:
            continue
        set_clauses.append(f"{k} = ?")
        values.append(v)
        if k in MANUAL_BILLING_FIELDS:
            manual_touch = True
    if not set_clauses:
        return
    set_clauses.append("updated_at = ?")
    values.append(ts)
    if manual_touch:
        set_clauses.append("last_manual_edit_at = ?")
        values.append(ts)
    values.append(record_id)
    with get_conn() as conn:
        conn.execute(f"UPDATE billing SET {', '.join(set_clauses)} WHERE id = ?", values)


def create_billing_manual(fields: dict):
    """Allow creating a brand-new billing record directly in the app."""
    ts = now()
    cols = [c for c in BILLING_IMPORT_FIELDS if c in fields]
    values = [fields[c] for c in cols]
    placeholders = ", ".join(["?"] * len(cols))
    with get_conn() as conn:
        conn.execute(
            f"""INSERT INTO billing ({", ".join(cols)}, created_at, updated_at, last_manual_edit_at)
                VALUES ({placeholders}, ?, ?, ?)""",
            values + [ts, ts, ts],
        )


def delete_billing(record_id: int):
    with get_conn() as conn:
        conn.execute("DELETE FROM billing WHERE id = ?", (record_id,))


# ---------------------------------------------------------------------------
# Master data: read
# ---------------------------------------------------------------------------

def get_all_master():
    with get_conn() as conn:
        rows = conn.execute("SELECT * FROM master_data ORDER BY uniq_id, period").fetchall()
        return [dict(r) for r in rows]


# ---------------------------------------------------------------------------
# Summary stats for dashboard
# ---------------------------------------------------------------------------

def get_summary():
    with get_conn() as conn:
        total_billing = conn.execute("SELECT COUNT(*) c FROM billing").fetchone()["c"]

        def status_counts(field):
            rows = conn.execute(
                f"SELECT COALESCE(NULLIF(TRIM({field}), ''), '(Blank)') AS k, COUNT(*) c "
                f"FROM billing GROUP BY k ORDER BY c DESC"
            ).fetchall()
            return {r["k"]: r["c"] for r in rows}

        isdp = status_counts("isdp_status")
        ibuy = status_counts("ibuy_status")
        overall_status = status_counts("status")
        cycle_counts = status_counts("cycle_month")
        total_master = conn.execute("SELECT COUNT(*) c FROM master_data").fetchone()["c"]

        return {
            "total_billing": total_billing,
            "isdp": isdp,
            "ibuy": ibuy,
            "status": overall_status,
            "cycle": cycle_counts,
            "total_master": total_master,
        }
