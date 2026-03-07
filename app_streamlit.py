"""
========================================================================
GST Reconciliation Tool – Streamlit Web Interface
========================================================================
Run with:
    streamlit run app.py

Requires:
    pip install streamlit pandas openpyxl numpy
========================================================================
"""

import io
import os
import sys
import tempfile

import pandas as pd
import streamlit as st

# Add parent dir to path if running from different location
sys.path.insert(0, os.path.dirname(__file__))

from gst_recon_engine import (
    read_file, detect_column_map, normalize_dataframe,
    GSTReconciler, build_summary,
)
from gst_recon_report import export_reconciliation_report


# ─────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="GST Reconciliation Tool",
    page_icon="🇮🇳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%);
        padding: 1.2rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .kpi-card {
        background: white;
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 4px solid;
        text-align: center;
    }
    .status-matched    { color: #065f46; background: #d1fae5; border-radius: 12px; padding: 2px 10px; font-weight: 700; font-size: 12px; }
    .status-mismatch   { color: #991b1b; background: #fee2e2; border-radius: 12px; padding: 2px 10px; font-weight: 700; font-size: 12px; }
    .status-only2b     { color: #78350f; background: #fef3c7; border-radius: 12px; padding: 2px 10px; font-weight: 700; font-size: 12px; }
    .status-onlybooks  { color: #4c1d95; background: #ede9fe; border-radius: 12px; padding: 2px 10px; font-weight: 700; font-size: 12px; }
    .status-probable   { color: #1e40af; background: #dbeafe; border-radius: 12px; padding: 2px 10px; font-weight: 700; font-size: 12px; }
    div[data-testid="metric-container"] { background: #f8fafc; border-radius: 8px; padding: 0.5rem; border: 1px solid #e2e8f0; }
    .footer-note { font-size: 11px; color: #94a3b8; margin-top: 1rem; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="main-header">
    <h2 style="margin:0;font-size:1.6rem;">🇮🇳 GST Reconciliation Tool</h2>
    <p style="margin:0;opacity:0.85;font-size:0.9rem;">
        GSTR-2B vs Purchase Register · B2B · Credit Notes · Debit Notes · Amendments
        · Rule 36(4) CGST Rules · Circular 183/15/2022-GST
    </p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────
# SIDEBAR – SETTINGS
# ─────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    tolerance = st.number_input(
        "Matching Tolerance (₹)",
        min_value=0.0, max_value=1000.0, value=10.0, step=1.0,
        help="Invoices with taxable value difference ≤ this amount are considered MATCHED"
    )
    date_window = st.number_input(
        "Probable Match Date Window (days)",
        min_value=0, max_value=30, value=5, step=1,
        help="±days for date-based probable matching"
    )
    period = st.text_input(
        "Reconciliation Period",
        placeholder="e.g. Apr-2024 to Mar-2025",
        help="Period label shown in the report"
    )
    st.markdown("---")
    st.markdown("**📌 Legal Reference**")
    st.markdown("""
    - CGST Act, 2017 – Sec 16
    - Rule 36(4) CGST Rules
    - Circular 183/15/2022-GST
    - GSTR-2B reconciliation mandate  
      (FY 2022-23 onwards)
    """)
    st.markdown("---")
    st.caption("Built for CA Advisory Firms · India")


# ─────────────────────────────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 📋 Upload GSTR-2B")
    gstr2b_file = st.file_uploader(
        "GSTR-2B Portal Download (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        key="gstr2b",
    )
    if gstr2b_file:
        st.success(f"✅ {gstr2b_file.name} uploaded")

with col2:
    st.markdown("#### 📚 Upload Purchase Register (Books)")
    books_file = st.file_uploader(
        "Purchase Register / ITC Register (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        key="books",
    )
    if books_file:
        st.success(f"✅ {books_file.name} uploaded")


# ─────────────────────────────────────────────────────────────────────
# TEMPLATE DOWNLOAD
# ─────────────────────────────────────────────────────────────────────

with st.expander("📥 Download Input Templates"):
    tc1, tc2 = st.columns(2)

    def make_template(is_gstr2b: bool) -> bytes:
        if is_gstr2b:
            data = {
                "GSTIN of Supplier": ["27AABCU9603R1ZX", "27AAACR5055K1Z5"],
                "Trade/Legal Name":  ["ABC Pvt Ltd", "XYZ Traders"],
                "Invoice Number":    ["INV-001", "CN-002"],
                "Invoice Date":      ["01-04-2024", "15-04-2024"],
                "Invoice Value":     [118000, 11800],
                "Taxable Value":     [100000, 10000],
                "Integrated Tax":    [18000, 1800],
                "Central Tax":       [0, 0],
                "State/UT Tax":      [0, 0],
                "Cess":              [0, 0],
                "Document Type":     ["Invoice", "Credit Note"],
            }
        else:
            data = {
                "Supplier GSTIN":  ["27AABCU9603R1ZX", "27AAACR5055K1Z5"],
                "Vendor Name":     ["ABC Pvt Ltd", "XYZ Traders"],
                "Invoice No":      ["INV-001", "CN-002"],
                "Invoice Date":    ["01-04-2024", "15-04-2024"],
                "Invoice Value":   [118000, 11800],
                "Taxable Value":   [100000, 10000],
                "IGST Amount":     [18000, 1800],
                "CGST Amount":     [0, 0],
                "SGST Amount":     [0, 0],
                "Cess":            [0, 0],
                "Type":            ["INV", "CN"],
            }
        df = pd.DataFrame(data)
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        return buf.getvalue()

    with tc1:
        st.download_button(
            "⬇ GSTR-2B Template", make_template(True),
            "GSTR2B_Template.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with tc2:
        st.download_button(
            "⬇ Purchase Register Template", make_template(False),
            "Purchase_Register_Template.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ─────────────────────────────────────────────────────────────────────
# RECONCILE BUTTON
# ─────────────────────────────────────────────────────────────────────

st.markdown("---")
run_btn = st.button(
    "▶ Run Reconciliation",
    disabled=(gstr2b_file is None or books_file is None),
    use_container_width=True,
    type="primary",
)

if not (gstr2b_file and books_file) and not run_btn:
    st.info("👆 Upload both GSTR-2B and Purchase Register files to proceed.")


# ─────────────────────────────────────────────────────────────────────
# RECONCILIATION LOGIC
# ─────────────────────────────────────────────────────────────────────

if run_btn and gstr2b_file and books_file:

    with st.spinner("🔄 Running reconciliation..."):

        # Save uploads to temp files
        with tempfile.NamedTemporaryFile(
            suffix=os.path.splitext(gstr2b_file.name)[1], delete=False
        ) as fg:
            fg.write(gstr2b_file.read())
            g_path = fg.name

        with tempfile.NamedTemporaryFile(
            suffix=os.path.splitext(books_file.name)[1], delete=False
        ) as fb:
            fb.write(books_file.read())
            b_path = fb.name

        try:
            raw_g = read_file(g_path)
            raw_b = read_file(b_path)

            g_map = detect_column_map(raw_g)
            b_map = detect_column_map(raw_b)

            gstr2b_norm = normalize_dataframe(raw_g, g_map, "GSTR2B")
            books_norm  = normalize_dataframe(raw_b, b_map, "BOOKS")

            gstr2b_norm = gstr2b_norm[gstr2b_norm["Invoice_Number_Clean"].str.strip() != ""].reset_index(drop=True)
            books_norm  = books_norm[books_norm["Invoice_Number_Clean"].str.strip() != ""].reset_index(drop=True)

            recon = GSTReconciler(gstr2b_norm, books_norm, tolerance=tolerance, probable_date_window=date_window)
            result_df = recon.reconcile()

            summary_data = build_summary(result_df, gstr2b_norm, books_norm)
            overall = summary_data["overall"]

        finally:
            os.unlink(g_path)
            os.unlink(b_path)

    st.success("✅ Reconciliation complete!")
    st.markdown("---")

    # ── KPI Cards ──
    st.markdown("### 📊 Summary")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("GSTR-2B ITC (₹)", f"{overall['Total as per GSTR-2B (₹)']:,.0f}")
    k2.metric("Books ITC (₹)",   f"{overall['Total as per Books (₹)']:,.0f}")
    k3.metric("Net Diff (₹)",    f"{overall['Net Difference (₹)']:,.0f}")
    k4.metric("✅ Matched",       overall["Matched"] + overall["CN Matched"] + overall["Amendment Match"])
    k5.metric("🔶 Only in 2B",   overall["Only in GSTR-2B (count)"])
    k6.metric("🔷 Only in Books",overall["Only in Books (count)"])

    # ── Category Table ──
    st.markdown("#### Category-wise Breakup")
    cat_df = summary_data["category"]
    st.dataframe(cat_df, use_container_width=True, hide_index=True)

    # ── Column Map Preview ──
    with st.expander("🔍 Detected Column Mapping"):
        mc1, mc2 = st.columns(2)
        with mc1:
            st.markdown("**GSTR-2B columns detected:**")
            st.json(g_map)
        with mc2:
            st.markdown("**Books columns detected:**")
            st.json(b_map)

    # ── Detail Table with Filters ──
    st.markdown("---")
    st.markdown("### 📄 Invoice-Level Reconciliation")

    f1, f2, f3 = st.columns(3)
    with f1:
        status_filter = st.multiselect(
            "Filter by Status",
            options=sorted(result_df["Status"].unique()),
            default=sorted(result_df["Status"].unique()),
        )
    with f2:
        doc_filter = st.multiselect(
            "Filter by Doc Type",
            options=sorted(result_df["Doc_Category"].unique()),
            default=sorted(result_df["Doc_Category"].unique()),
        )
    with f3:
        gstin_search = st.text_input("Search GSTIN / Invoice No")

    filtered = result_df[
        result_df["Status"].isin(status_filter) &
        result_df["Doc_Category"].isin(doc_filter)
    ]
    if gstin_search:
        filtered = filtered[
            filtered["GSTIN"].str.contains(gstin_search, case=False, na=False) |
            filtered["Invoice_Number"].str.contains(gstin_search, case=False, na=False)
        ]

    display_cols = [
        "GSTIN", "Supplier_Name", "Invoice_Number", "Invoice_Date", "Doc_Category",
        "GSTR2B_Taxable", "GSTR2B_Total_Tax",
        "Books_Taxable", "Books_Total_Tax",
        "Difference_Taxable", "Difference_Tax", "Status"
    ]
    st.dataframe(
        filtered[[c for c in display_cols if c in filtered.columns]],
        use_container_width=True,
        hide_index=True,
        height=420,
    )
    st.caption(f"Showing {len(filtered):,} of {len(result_df):,} records")

    # ── Vendor Summary ──
    st.markdown("---")
    st.markdown("### 🏢 Vendor-wise ITC Summary")
    from gst_recon_engine import build_vendor_summary
    vendor_df = build_vendor_summary(result_df)
    st.dataframe(vendor_df, use_container_width=True, hide_index=True)

    # ── Export ──
    st.markdown("---")
    st.markdown("### ⬇ Download Report")

    buf = io.BytesIO()
    export_reconciliation_report(
        result_df, summary_data, gstr2b_norm, books_norm,
        output_path=buf,
        period=period,
        tolerance=tolerance,
    )
    buf.seek(0)

    st.download_button(
        label="📥 Download Excel Reconciliation Report",
        data=buf.getvalue(),
        file_name=f"GST_Recon_{period.replace(' ','_') if period else 'Report'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    st.markdown("""
    <div class="footer-note">
    ⚖️ <strong>Legal References:</strong> CGST Act 2017 §16 | Rule 36(4) CGST Rules | Circular 183/15/2022-GST |
    ITC eligible only if reflected in GSTR-2B per amendment w.e.f. 01-Jan-2022.<br>
    This tool is for professional use by Chartered Accountants. Results should be reviewed before filing.
    </div>
    """, unsafe_allow_html=True)
