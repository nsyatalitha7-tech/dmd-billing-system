"""
DMD Billing Management System
==============================
Streamlit + SQLite internal dashboard that replaces manual Excel updates
for daily ISDP Status / IBuy Status / Remarks tracking.

Run with:  streamlit run app.py
"""

import tempfile
from datetime import datetime

import pandas as pd
import streamlit as st

import db
import importer

st.set_page_config(page_title="DMD Billing Management System", layout="wide")

db.init_db()

CYCLE_SORT_KEY = lambda c: int(c[1:]) if c and c[1:].isdigit() else 999  # "M4" -> 4

CSS = """
<style>
.block-container {padding-top: 2rem;}
div[data-testid="stMetric"] {
    background: #F7F7F5;
    border: 1px solid #E4E2DD;
    border-radius: 8px;
    padding: 14px 16px 10px 16px;
}
.badge {
    display:inline-block; padding:2px 10px; border-radius:12px;
    font-size:0.8rem; font-weight:600;
}
.badge-approved, .badge-invoice { background:#DCEFE1; color:#1E7B41; }
.badge-hold, .badge-pending, .badge-blank { background:#F1EFC9; color:#8A7300; }
.badge-eror, .badge-reject, .badge-cancelinvoice { background:#F6D9D6; color:#B3261E; }
.badge-prepare, .badge-other { background:#E4E2DD; color:#55524B; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def status_class(status: str) -> str:
    if not status:
        return "badge-blank"
    s = status.strip().lower().replace(" ", "")
    known = ["approved", "invoice", "hold", "pending", "eror", "reject",
             "cancelinvoice", "prepare", "blank"]
    return f"badge-{s}" if s in known else "badge-other"


def badge(status: str) -> str:
    label = status if status else "(Blank)"
    return f'<span class="badge {status_class(status)}">{label}</span>'


# ---------------------------------------------------------------------------
# Sidebar navigation
# ---------------------------------------------------------------------------

st.sidebar.title("DMD Billing")
page = st.sidebar.radio(
    "Navigate",
    ["Dashboard", "Billing", "Master Data", "Import Data"],
    label_visibility="collapsed",
)

summary = db.get_summary()
st.sidebar.markdown("---")
st.sidebar.caption(f"Billing records: **{summary['total_billing']}**")
st.sidebar.caption(f"Master records: **{summary['total_master']}**")
st.sidebar.caption(f"DB file: `{db.DB_PATH.name}`")


# ---------------------------------------------------------------------------
# DASHBOARD
# ---------------------------------------------------------------------------

if page == "Dashboard":
    st.title("Dashboard")

    s = db.get_summary()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Billing Records", s["total_billing"])
    c2.metric("ISDP Approved", s["isdp"].get("Approved", 0))
    c3.metric("ISDP Hold", s["isdp"].get("Hold", 0))
    c4.metric("ISDP Eror / Reject", s["isdp"].get("Eror", 0) + s["isdp"].get("Reject", 0))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("IBuy Invoice", s["ibuy"].get("Invoice", 0))
    c6.metric("IBuy Cancelled", s["ibuy"].get("Cancel invoice", 0))
    c7.metric("Awaiting ISDP", s["isdp"].get("(Blank)", 0))
    c8.metric("Awaiting IBuy", s["ibuy"].get("(Blank)", 0))

    st.markdown("### Breakdown")
    bcol1, bcol2, bcol3 = st.columns(3)

    with bcol1:
        st.markdown("**ISDP Status**")
        df = pd.DataFrame(list(s["isdp"].items()), columns=["Status", "Count"]).sort_values(
            "Count", ascending=False
        )
        st.dataframe(df, hide_index=True, use_container_width=True)

    with bcol2:
        st.markdown("**IBuy Status**")
        df = pd.DataFrame(list(s["ibuy"].items()), columns=["Status", "Count"]).sort_values(
            "Count", ascending=False
        )
        st.dataframe(df, hide_index=True, use_container_width=True)

    with bcol3:
        st.markdown("**By Cycle Month**")
        items = sorted(s["cycle"].items(), key=lambda kv: CYCLE_SORT_KEY(kv[0]))
        df = pd.DataFrame(items, columns=["Cycle", "Count"])
        st.dataframe(df, hide_index=True, use_container_width=True)

    st.markdown("### Overall Billing Status")
    df = pd.DataFrame(list(s["status"].items()), columns=["Status", "Count"]).sort_values(
        "Count", ascending=False
    )
    st.bar_chart(df.set_index("Status"))

    with st.expander("Recent import history"):
        log = db.get_import_log(20)
        if log:
            st.dataframe(pd.DataFrame(log), hide_index=True, use_container_width=True)
        else:
            st.caption("No imports yet.")


# ---------------------------------------------------------------------------
# BILLING
# ---------------------------------------------------------------------------

elif page == "Billing":
    st.title("Billing")

    all_rows = db.get_all_billing()
    df = pd.DataFrame(all_rows)

    if df.empty:
        st.info("No billing records yet. Go to **Import Data** to load the Excel file.")
    else:
        with st.expander("Filters", expanded=True):
            f1, f2, f3, f4 = st.columns(4)
            search = f1.text_input("Search Site Code / PO")
            cycles = sorted(df["cycle_month"].dropna().unique(), key=CYCLE_SORT_KEY)
            cycle_filter = f2.multiselect("Cycle Month", cycles)
            isdp_opts = sorted([x for x in df["isdp_status"].dropna().unique()])
            isdp_filter = f3.multiselect("ISDP Status", isdp_opts)
            ibuy_opts = sorted([x for x in df["ibuy_status"].dropna().unique()])
            ibuy_filter = f4.multiselect("IBuy Status", ibuy_opts)

        fdf = df.copy()
        if search:
            mask = (
                fdf["site_code"].fillna("").str.contains(search, case=False)
                | fdf["po"].fillna("").str.contains(search, case=False)
            )
            fdf = fdf[mask]
        if cycle_filter:
            fdf = fdf[fdf["cycle_month"].isin(cycle_filter)]
        if isdp_filter:
            fdf = fdf[fdf["isdp_status"].isin(isdp_filter)]
        if ibuy_filter:
            fdf = fdf[fdf["ibuy_status"].isin(ibuy_filter)]

        st.caption(f"{len(fdf)} of {len(df)} records")

        display_cols = [
            "id", "site_code", "cycle_month", "po", "site_name", "regional",
            "status", "isdp_status", "ibuy_status", "remarks", "total_price",
        ]
        display_cols = [c for c in display_cols if c in fdf.columns]
        st.dataframe(
            fdf[display_cols].sort_values(["site_code", "cycle_month"]),
            hide_index=True,
            use_container_width=True,
            height=420,
        )

        st.markdown("### Edit a record")
        options = {
            f"{r['site_code']} | {r['cycle_month']} | PO {r['po'] or '-'} (id {r['id']})": r["id"]
            for _, r in fdf.sort_values(["site_code", "cycle_month"]).iterrows()
        }
        if options:
            choice = st.selectbox("Select a billing record to edit", list(options.keys()))
            record_id = options[choice]
            record = db.get_billing_by_id(record_id)

            with st.form("edit_billing_form"):
                col1, col2 = st.columns(2)
                site_code = col1.text_input("Site Code", value=record["site_code"], disabled=True)
                cycle_month = col2.text_input("Cycle Month", value=record["cycle_month"], disabled=True)
                po = st.text_input("PO", value=record.get("po") or "")

                col3, col4 = st.columns(2)
                isdp_status = col3.text_input("ISDP Status", value=record.get("isdp_status") or "")
                ibuy_status = col4.text_input("IBuy Status", value=record.get("ibuy_status") or "")
                remarks = st.text_area("Remarks", value=record.get("remarks") or "")

                saved = st.form_submit_button("Save", type="primary")
                if saved:
                    db.update_billing_full(
                        record_id,
                        {
                            "po": po.strip() or None,
                            "isdp_status": isdp_status.strip() or None,
                            "ibuy_status": ibuy_status.strip() or None,
                            "remarks": remarks.strip() or None,
                        },
                    )
                    st.success(
                        f"Saved. {site_code} | {cycle_month} updated at "
                        f"{datetime.now().strftime('%H:%M:%S')}."
                    )
                    st.rerun()

            with st.expander("Delete this record"):
                st.warning("This permanently removes the billing record from the database.")
                if st.button("Delete record", key=f"del_{record_id}"):
                    db.delete_billing(record_id)
                    st.success("Deleted.")
                    st.rerun()

    st.markdown("---")
    with st.expander("Add a new billing record manually"):
        with st.form("new_billing_form"):
            n1, n2, n3 = st.columns(3)
            new_site = n1.text_input("Site Code *")
            new_cycle = n2.text_input("Cycle Month * (e.g. M7)")
            new_po = n3.text_input("PO")
            n4, n5 = st.columns(2)
            new_isdp = n4.text_input("ISDP Status")
            new_ibuy = n5.text_input("IBuy Status")
            new_remarks = st.text_area("Remarks")
            create = st.form_submit_button("Create record")
            if create:
                if not new_site.strip() or not new_cycle.strip():
                    st.error("Site Code and Cycle Month are required.")
                else:
                    existing = [
                        r for r in db.get_all_billing()
                        if r["site_code"] == new_site.strip() and r["cycle_month"] == new_cycle.strip()
                    ]
                    if existing:
                        st.error(
                            f"A record for {new_site.strip()} | {new_cycle.strip()} already exists "
                            f"(id {existing[0]['id']}). Edit it above instead."
                        )
                    else:
                        db.create_billing_manual({
                            "site_code": new_site.strip(),
                            "cycle_month": new_cycle.strip(),
                            "po": new_po.strip() or None,
                            "isdp_status": new_isdp.strip() or None,
                            "ibuy_status": new_ibuy.strip() or None,
                            "remarks": new_remarks.strip() or None,
                        })
                        st.success("Created.")
                        st.rerun()


# ---------------------------------------------------------------------------
# MASTER DATA
# ---------------------------------------------------------------------------

elif page == "Master Data":
    st.title("Master Data")
    st.caption(
        "Read-only reference data (Uniq-ID, Account, Region, Period, Productivity, "
        "Payment Coef.). Use **Import Data** to sync it from a new Excel snapshot."
    )

    rows = db.get_all_master()
    df = pd.DataFrame(rows)

    if df.empty:
        st.info("No master data yet. Go to **Import Data** to sync it from Excel.")
    else:
        f1, f2, f3 = st.columns(3)
        search = f1.text_input("Search Uniq-ID / Name")
        region_opts = sorted([x for x in df["region"].dropna().unique()])
        region_filter = f2.multiselect("Region", region_opts)
        period_opts = sorted(df["period"].dropna().unique(), key=CYCLE_SORT_KEY)
        period_filter = f3.multiselect("Period", period_opts)

        fdf = df.copy()
        if search:
            mask = (
                fdf["uniq_id"].fillna("").str.contains(search, case=False)
                | fdf["name_id"].fillna("").str.contains(search, case=False)
            )
            fdf = fdf[mask]
        if region_filter:
            fdf = fdf[fdf["region"].isin(region_filter)]
        if period_filter:
            fdf = fdf[fdf["period"].isin(period_filter)]

        st.caption(f"{len(fdf)} of {len(df)} records")

        display_cols = [
            "uniq_id", "column1", "name_id", "account", "region", "period",
            "period_date", "period_status", "productivity_range", "productivity",
            "payment_coef",
        ]
        display_cols = [c for c in display_cols if c in fdf.columns]
        st.dataframe(
            fdf[display_cols].sort_values(["uniq_id", "period"]),
            hide_index=True,
            use_container_width=True,
            height=500,
        )


# ---------------------------------------------------------------------------
# IMPORT DATA
# ---------------------------------------------------------------------------

elif page == "Import Data":
    st.title("Import Data")
    st.write(
        "Upload the DMD Excel workbook to do the initial import or a periodic "
        "master-data sync. Daily ISDP / IBuy / Remarks edits made in the app are "
        "**preserved by default** — Excel only fills in blanks and adds new records "
        "unless you explicitly opt in below."
    )

    uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])

    if uploaded is not None:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded.getbuffer())
            tmp_path = tmp.name

        sheets = importer.list_sheets(tmp_path)

        tab1, tab2 = st.tabs(["Billing Import", "Master Data Sync"])

        # --- Billing import ---
        with tab1:
            default_sheet = importer.guess_billing_sheet(sheets)
            billing_sheet = st.selectbox(
                "Billing / Tagihan sheet", sheets,
                index=sheets.index(default_sheet), key="billing_sheet"
            )
            overwrite_manual = st.checkbox(
                "Also overwrite ISDP Status / IBuy Status / Remarks from this Excel "
                "(off by default — leave off for routine imports)",
                value=False,
            )

            if st.button("Preview billing import"):
                try:
                    rows, dup_report, skipped_no_key = importer.read_billing_sheet(
                        tmp_path, billing_sheet
                    )
                    st.session_state["billing_preview"] = {
                        "rows": rows, "dup_report": dup_report,
                        "skipped_no_key": skipped_no_key, "sheet": billing_sheet,
                        "filename": uploaded.name,
                    }
                except ValueError as e:
                    st.error(str(e))
                    st.session_state.pop("billing_preview", None)

            preview = st.session_state.get("billing_preview")
            if preview and preview["sheet"] == billing_sheet:
                rows = preview["rows"]
                dup_report = preview["dup_report"]
                st.write(
                    f"Found **{len(rows)}** billing records "
                    f"({len(dup_report)} duplicate Site Code + Cycle Month rows resolved "
                    f"by keeping the most recent, {preview['skipped_no_key']} rows skipped "
                    f"for missing Site Code/Cycle Month)."
                )
                if dup_report:
                    with st.expander(f"View {len(dup_report)} resolved duplicates"):
                        st.dataframe(pd.DataFrame(dup_report), hide_index=True, use_container_width=True)

                existing_keys = {
                    (r["site_code"], r["cycle_month"]) for r in db.get_all_billing()
                }
                new_count = sum(
                    1 for r in rows if (r["site_code"], r["cycle_month"]) not in existing_keys
                )
                update_count = len(rows) - new_count
                st.write(f"This will **create {new_count}** new records and **update {update_count}** existing records.")
                if overwrite_manual:
                    st.warning(
                        "Manual field overwrite is ON: ISDP Status, IBuy Status and Remarks "
                        "for existing records WILL be replaced with the Excel values."
                    )

                if st.button("Confirm & Import Billing Data", type="primary"):
                    result = db.import_billing_rows(rows, overwrite_manual=overwrite_manual)
                    db.log_import(
                        "billing", preview["filename"], billing_sheet,
                        result["inserted"], result["updated"], len(dup_report),
                        overwrite_manual,
                        detail=f"skipped_no_key={preview['skipped_no_key']}",
                    )
                    st.success(
                        f"Import complete: {result['inserted']} inserted, "
                        f"{result['updated']} updated."
                    )
                    del st.session_state["billing_preview"]
                    st.rerun()

        # --- Master data sync ---
        with tab2:
            default_master = importer.guess_master_sheet(sheets)
            master_sheet = st.selectbox(
                "Master Data sheet", sheets,
                index=sheets.index(default_master), key="master_sheet"
            )

            if st.button("Preview master data sync"):
                try:
                    rows = importer.read_master_sheet(tmp_path, master_sheet)
                    st.session_state["master_preview"] = {
                        "rows": rows, "sheet": master_sheet, "filename": uploaded.name,
                    }
                except ValueError as e:
                    st.error(str(e))
                    st.session_state.pop("master_preview", None)

            preview = st.session_state.get("master_preview")
            if preview and preview["sheet"] == master_sheet:
                rows = preview["rows"]
                st.write(f"Found **{len(rows)}** unique Uniq-ID + Period master records.")
                st.dataframe(pd.DataFrame(rows).head(20), hide_index=True, use_container_width=True)

                existing_keys = {r["column1"] for r in db.get_all_master()}
                new_count = sum(1 for r in rows if r["column1"] not in existing_keys)
                update_count = len(rows) - new_count
                st.write(f"This will **create {new_count}** new records and **update {update_count}** existing records.")

                if st.button("Confirm & Sync Master Data", type="primary"):
                    result = db.import_master_rows(rows)
                    db.log_import(
                        "master", preview["filename"], master_sheet,
                        result["inserted"], result["updated"], 0, False,
                    )
                    st.success(
                        f"Sync complete: {result['inserted']} inserted, "
                        f"{result['updated']} updated."
                    )
                    del st.session_state["master_preview"]
                    st.rerun()
