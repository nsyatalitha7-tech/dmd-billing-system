# DMD Billing Management System

A small internal dashboard that replaces manual Excel updates for daily
billing status tracking. Excel is used only for the initial import and
occasional master-data sync — day-to-day edits happen in the app and are
stored in a real database (SQLite).

## What was found in your Excel file

`Tagihan_DMD_2209.xlsx` contains 9 sheets. The app uses two of them:

| Sheet | Role | Notes |
|---|---|---|
| `Master Data 14-9-2026` (or `Master Data 9-9-26`) | Master Data | Header is on row 1. Only the 11 fields you specified are imported: Uniq-ID, Column1 (aliased from `Kolom1` on the older sheet), Name ID, Account, Region, Period, Period Date, Period Status, Productivity Range, Productivity, Payment Coef. Each Uniq-ID+Period repeats several times (once per Document Category/Type) — the importer de-duplicates on `Column1` since those required fields are identical across the repeats. |
| `Tagihan` | Billing | Real header is on **row 2**, not row 1 (row 1 has a stray value). The importer auto-detects the header row instead of assuming a fixed position. Key fields: **Site Code** + **Period Tagih** (→ Cycle Month) as the unique billing key, plus PO NO., ISDP Status, Ibuy Status, Remarks, and other billing detail fields (Site Name, Regional, Item Description, Unit Price, Total Price, Acceptance Date, Status, Invoice info, etc). |

Status values actually found in the data (so the dashboard summarizes real
categories, not assumed ones):
- **ISDP Status:** Approved, Hold, Eror, Prepare, Reject, (blank)
- **IBuy Status:** Invoice, Cancel invoice, (blank)

**Data-quality notes surfaced during import:**
- ~38 rows in `Tagihan` share the same Site Code + Cycle Month (e.g. a PO was
  re-issued/corrected). The importer keeps the row with the latest Acceptance
  Date and reports every row it skipped in the Import Data screen, so nothing
  is silently lost — you can review the list before confirming.
- Some text values had trailing spaces (`"Reject "` vs `"Reject"`) which are
  stripped so filters and dashboard counts aren't fragmented.
- ~200 rows in `Tagihan` have no Site Code or Cycle Month (blank/footer rows)
  and are skipped automatically.

## How the database works

- **`billing`** table — one row per `(site_code, cycle_month)`, enforced with
  a UNIQUE constraint so re-importing the same Excel can never create
  duplicates.
- **`master_data`** table — one row per `column1` (Uniq-ID + Period).
- **`import_log`** table — history of every import/sync run.

### The core rule: manual edits win

`isdp_status`, `ibuy_status`, and `remarks` are treated as **manually
maintained** fields. Excel imports only fill them in for **brand-new**
records. For records that already exist, the importer leaves those three
fields untouched **unless you explicitly tick "Also overwrite ISDP Status /
IBuy Status / Remarks from this Excel"** on the Import Data page. Every other
billing field (PO, prices, dates, etc.) is refreshed from Excel on every
import, but only when the Excel value is non-blank — a blank cell never wipes
out data you already have.

Master data sync always updates the required master fields on the matching
`column1`, since master data isn't something you edit by hand in the app.

## Setup

```bash
cd dmd_billing
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Open the URL Streamlit prints (usually http://localhost:8501).

## First-time use

1. Go to **Import Data**, upload `Tagihan_DMD_2209.xlsx`.
2. On the **Master Data Sync** tab, pick the newest Master Data sheet,
   preview, and confirm — this populates the Master Data page.
3. On the **Billing Import** tab, pick the `Tagihan` sheet, preview (review
   the duplicate/skip report), and confirm — this populates the Billing page
   and Dashboard.
4. From then on, do daily status updates on the **Billing** page (select a
   record → edit ISDP/IBuy/Remarks → Save). Only re-run Import Data when you
   get a new Excel snapshot (e.g. new sites added, or a periodic master-data
   refresh) — leave "overwrite manual fields" unchecked for routine re-imports.

## Project structure

```
dmd_billing/
  app.py         Streamlit UI: Dashboard, Billing, Master Data, Import Data
  db.py          SQLite schema + all read/write/upsert logic
  importer.py    Excel reading, header-detection, column normalization, dedup
  requirements.txt
  data/
    dmd_billing.db   created automatically on first run
```

## Notes / things to keep in mind

- The database file `data/dmd_billing.db` is the system of record — back it
  up periodically (it's a single portable file).
- If a future Excel export adds/removes columns in Master Data, nothing
  breaks: only the 11 required fields are read, matched by name (with a
  couple of known aliases like `Column1`/`Kolom1`); unrecognized columns are
  ignored.
- If the Tagihan sheet's column names change, the importer will raise a
  clear error naming which expected columns it couldn't find, rather than
  importing garbage.
