"""
importer.py
-----------
Reads the DMD Excel workbook and turns it into clean rows ready for db.py.

Handles the real-world messiness found in the source file:
- The "Tagihan" (billing) sheet has its real header on the SECOND row
  (row 1 is a stray timestamp / merged-cell artifact), while Master Data
  sheets have their header on the first row. We auto-detect the header
  row instead of assuming a fixed position.
- Master Data sheets across snapshots use slightly different column
  names for the same field (e.g. "Column1" vs "Kolom1"). We match on a
  normalized (lowercased, punctuation-stripped) column name and a small
  alias table.
- Text values have stray whitespace ("Reject " vs "Reject") which is
  stripped so status filters/summaries aren't fragmented.
- The same (Site Code, Cycle Month) combination can legitimately appear
  more than once in a raw export (e.g. a corrected/re-issued PO). Since
  the system's data model requires this combination to be unique, the
  importer keeps the most recent row (by Acceptance Date, falling back
  to sheet order) and reports every row it had to skip so nothing is
  silently lost.
"""

import re
from datetime import datetime

import pandas as pd
import openpyxl


# ---------------------------------------------------------------------------
# Column normalization
# ---------------------------------------------------------------------------

def _norm(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"[^a-z0-9]+", "", s)
    return s


MASTER_ALIASES = {
    "uniqid": "uniq_id",
    "column1": "column1",
    "kolom1": "column1",
    "nameid": "name_id",
    "account": "account",
    "region": "region",
    "period": "period",
    "perioddate": "period_date",
    "periodstatus": "period_status",
    "productivityrange": "productivity_range",
    "productivity": "productivity",
    "paymentcoef": "payment_coef",
}

BILLING_ALIASES = {
    "sitecode": "site_code",
    "name": "name",
    "projectbefore": "project_before",
    "projectassign": "project_assign",
    "newproject": "new_project",
    "periodtagih": "cycle_month",
    "joindate": "join_date",
    "joinmonth": "join_month",
    "sitename": "site_name",
    "regional": "regional",
    "siteid": "site_id_ext",
    "subcontractno": "sub_contract_no",
    "prno": "pr_no",
    "pono": "po",
    "polineno": "po_line_no",
    "shipmentno": "shipment_no",
    "itemcode": "item_code",
    "itemdescription": "item_description",
    "unitprice": "unit_price",
    "requestedqty": "requested_qty",
    "paymentcoef": "payment_coef",
    "totalprice": "total_price",
    "acceptancedate": "acceptance_date",
    "status": "status",
    "isdpstatus": "isdp_status",
    "remarks": "remarks",
    "action": "action",
    "ibuystatus": "ibuy_status",
    "idproject": "id_project",
    "sonumber": "so_number",
    "invoicenumber": "invoice_number",
    "invoicedate": "invoice_date",
    "invoicemonth": "invoice_month",
}

REQUIRED_MASTER_TARGETS = [
    "uniq_id", "column1", "name_id", "account", "region", "period",
    "period_date", "period_status", "productivity_range", "productivity",
    "payment_coef",
]

REQUIRED_BILLING_TARGETS = ["site_code", "cycle_month"]


def list_sheets(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    return wb.sheetnames


def _find_header_row(path, sheet, expected_norm_names, max_scan=5):
    """Scan the first `max_scan` rows for the one that best matches the
    expected (normalized) column names, and return its 0-based index."""
    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=max_scan)
    best_row, best_score = 0, -1
    for i in range(min(max_scan, len(raw))):
        row_vals = [_norm(v) for v in raw.iloc[i].tolist()]
        score = sum(1 for v in row_vals if v in expected_norm_names)
        if score > best_score:
            best_score = score
            best_row = i
    return best_row


def _rename_columns(df, alias_map):
    new_cols = {}
    for col in df.columns:
        key = _norm(col)
        if key in alias_map:
            new_cols[col] = alias_map[key]
    df = df.rename(columns=new_cols)
    return df


def _is_na(v):
    if v is None:
        return True
    try:
        result = pd.isna(v)
        return bool(result)
    except (TypeError, ValueError):
        return False


def _clean_str(v):
    if _is_na(v):
        return None
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "nat", "none"):
        return None
    return s


def _clean_id_str(v):
    """Like _clean_str, but repairs IDs that pandas turned into float/
    scientific-notation (e.g. a numeric PR/PO/Item code -> 1.234e+17)."""
    if _is_na(v):
        return None
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return repr(v)
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "nat", "none"):
        return None
    return s


def _clean_num(v):
    if _is_na(v):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _clean_date(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return None
    return s


# ---------------------------------------------------------------------------
# Master data
# ---------------------------------------------------------------------------

def read_master_sheet(path, sheet):
    expected = set(MASTER_ALIASES.keys())
    header_row = _find_header_row(path, sheet, expected)
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df = _rename_columns(df, MASTER_ALIASES)

    missing = [c for c in ["uniq_id", "column1"] if c not in df.columns]
    if missing:
        raise ValueError(
            f"Sheet '{sheet}' doesn't look like Master Data (missing {missing}). "
            f"Pick a different sheet."
        )

    rows = []
    for _, r in df.iterrows():
        row = {}
        for target in REQUIRED_MASTER_TARGETS:
            val = r[target] if target in df.columns else None
            if target == "payment_coef":
                row[target] = _clean_num(val)
            elif target == "period_date":
                row[target] = _clean_str(val)  # kept as text range "YYYY-MM-DD - YYYY-MM-DD"
            else:
                row[target] = _clean_str(val)
        if not row["uniq_id"] or not row["column1"]:
            continue
        rows.append(row)

    # De-duplicate on column1 (a Master Data sheet repeats the same
    # Uniq-ID/Period combo once per Document Category/Type row; the
    # required master fields are identical across those repeats).
    dedup = {}
    for row in rows:
        dedup[row["column1"]] = row  # last one wins, values are identical anyway
    return list(dedup.values())


# ---------------------------------------------------------------------------
# Billing data
# ---------------------------------------------------------------------------

def read_billing_sheet(path, sheet):
    expected = set(BILLING_ALIASES.keys())
    header_row = _find_header_row(path, sheet, expected)
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df = _rename_columns(df, BILLING_ALIASES)

    missing = [c for c in REQUIRED_BILLING_TARGETS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Sheet '{sheet}' doesn't look like the Billing/Tagihan sheet "
            f"(missing {missing}). Pick a different sheet."
        )

    text_fields = [
        "site_code", "cycle_month", "name", "project_assign", "site_name",
        "regional", "sub_contract_no", "status", "isdp_status", "remarks",
        "ibuy_status", "id_project", "so_number", "invoice_number", "invoice_month",
    ]
    id_fields = ["po", "pr_no", "po_line_no", "shipment_no", "item_code", "item_description"]
    num_fields = ["unit_price", "requested_qty", "payment_coef", "total_price"]
    date_fields = ["acceptance_date", "invoice_date"]

    rows = []
    skipped_no_key = 0
    for _, r in df.iterrows():
        row = {}
        for f in text_fields:
            row[f] = _clean_str(r[f]) if f in df.columns else None
        for f in id_fields:
            row[f] = _clean_id_str(r[f]) if f in df.columns else None
        for f in num_fields:
            row[f] = _clean_num(r[f]) if f in df.columns else None
        for f in date_fields:
            row[f] = _clean_date(r[f]) if f in df.columns else None

        if not row.get("site_code") or not row.get("cycle_month"):
            skipped_no_key += 1
            continue
        rows.append(row)

    # Resolve in-batch duplicates on (site_code, cycle_month): keep the
    # row with the latest Acceptance Date, falling back to the last row
    # encountered (bottom-most = most recently added in the sheet).
    best = {}
    duplicates_report = []
    for row in rows:
        key = (row["site_code"], row["cycle_month"])
        if key not in best:
            best[key] = row
            continue
        existing = best[key]
        existing_date = existing.get("acceptance_date") or ""
        new_date = row.get("acceptance_date") or ""
        winner, loser = (row, existing) if new_date >= existing_date else (existing, row)
        duplicates_report.append({
            "site_code": key[0],
            "cycle_month": key[1],
            "kept_po": winner.get("po"),
            "skipped_po": loser.get("po"),
        })
        best[key] = winner

    return list(best.values()), duplicates_report, skipped_no_key


def guess_master_sheet(sheets):
    candidates = [s for s in sheets if "master" in s.lower()]
    if not candidates:
        return sheets[0]
    # Prefer the most recently dated-looking sheet name (simple heuristic:
    # the last one alphabetically / in the workbook tends to be newest).
    return candidates[-1]


def guess_billing_sheet(sheets):
    for s in sheets:
        if "tagihan" in s.lower() or "billing" in s.lower():
            return s
    return sheets[0]
