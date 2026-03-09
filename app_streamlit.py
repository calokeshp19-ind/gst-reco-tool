"""
GST Reconciliation Tool — Streamlit UI v2
Supports GSTR-2A and GSTR-2B automatically.
"""

import io
import tempfile
import os
import numpy as np
import pandas as pd
import streamlit as st

from gst_recon_engine import (
    detect_file_type, read_gst_file, read_books,
    reconcile, build_summary, build_vendor_summary,
    TOLERANCE, DATE_WINDOW,
)
from gst_recon_report import write_report

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GST Reconciliation Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1E3A5F, #2563EB);
        color: white; padding: 20px 28px; border-radius: 12px;
        margin-bottom: 24px;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .main-header p  { color: #CBD5E1; margin: 4px 0 0 0; font-size: 0.9rem; }
    .metric-card {
        background: white; padding: 16px; border-radius: 10px;
        border-left: 4px solid #2563EB;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .status-matched  { background:#D1FAE5; padding:3px 10px; border-radius:12px; color:#065F46; font-weight:600; font-size:0.82rem; }
    .status-probable { background:#DBEAFE; padding:3px 10px; border-radius:12px; color:#1E40AF; font-weight:600; font-size:0.82rem; }
    .status-mismatch { background:#FEE2E2; padding:3px 10px; border-radius:12px; color:#7F1D1D; font-weight:600; font-size:0.82rem; }
    .status-only2a   { background:#FEF3C7; padding:3px 10px; border-radius:12px; color:#78350F; font-weight:600; font-size:0.82rem; }
    .status-onlybks  { background:#EDE9FE; padding:3px 10px; border-radius:12px; color:#4C1D95; font-weight:600; font-size:0.82rem; }
    .info-box {
        background:#EFF6FF; border:1px solid #BFDBFE; padding:12px 16px;
        border-radius:8px; margin:8px 0; font-size:0.88rem; color:#1E40AF;
    }
    .warn-box {
        background:#FFFBEB; border:1px solid #FDE68A; padding:12px 16px;
        border-radius:8px; margin:8px 0; font-size:0.88rem; color:#78350F;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 GST Reconciliation Tool</h1>
    <p>GSTR-2A / GSTR-2B vs Books of Accounts &nbsp;|&nbsp; Rule 36(4) CGST Rules &nbsp;|&nbsp; Powered by Python</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    tolerance   = st.number_input("Tolerance (Rs)", 0.0, 1000.0, float(TOLERANCE), 1.0,
                                  help="Max tax difference allowed for a MATCHED status")
    date_window = st.number_input("Date Window (days)", 0, 30, DATE_WINDOW, 1,
                                  help="Max date difference for Level 1 matching")
    period      = st.text_input("Period", "Apr-2024 to Mar-2025",
                                help="Reconciliation period for report header")

    st.markdown("---")
    st.markdown("### 📋 Supported Formats")
    st.markdown("""
    **GST Statement:**
    - ✅ GSTR-2A (Portal export)
    - ✅ GSTR-2B (Portal export)

    **Books of Accounts:**
    - ✅ Tally export
    - ✅ Any Excel with GSTIN, IGST, CGST, SGST columns

    **Format auto-detected on upload**
    """)

    st.markdown("---")
    st.markdown("### 📌 Matching Logic")
    st.markdown(f"""
    | Level | Logic | Status |
    |---|---|---|
    | L1 | GSTIN + Tax + Date ±{date_window}d | MATCHED |
    | L2 | GSTIN + Tax | PROBABLE |
    | L3 | GSTIN only | MISMATCH |
    | L4 | No match | ONLY IN 2A/BOOKS |
    """)

    st.markdown("---")
    st.markdown("### ⚖️ Legal Note")
    st.markdown("""
    <div class="warn-box">
    ITC eligibility must be verified against <b>GSTR-2B</b> per Rule 36(4) CGST Rules.
    GSTR-2A is dynamic and for reference only.
    </div>
    """, unsafe_allow_html=True)

# ── File Upload ───────────────────────────────────────────────────────────────
st.markdown("### 📂 Upload Files")
col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 📄 GST Statement (2A or 2B)")
    gst_file = st.file_uploader(
        "Upload GSTR-2A or GSTR-2B Excel file",
        type=["xlsx","xls"], key="gst_file"
    )
    if gst_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(gst_file.read())
            gst_path = tmp.name
        detected = detect_file_type(gst_path)
        label    = "GSTR-2A" if detected=="2A" else "GSTR-2B" if detected=="2B" else "Unknown"
        color    = "green" if detected in ["2A","2B"] else "red"
        st.markdown(f"**Auto-detected:** :{color}[{label}]")
        st.markdown(f"""<div class="info-box">✅ <b>{gst_file.name}</b> uploaded successfully as <b>{label}</b></div>""",
                    unsafe_allow_html=True)

with col2:
    st.markdown("#### 📒 Books of Accounts")
    books_file = st.file_uploader(
        "Upload Books Excel file (Tally export or similar)",
        type=["xlsx","xls"], key="books_file"
    )
    if books_file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(books_file.read())
            books_path = tmp.name
        st.markdown(f"""<div class="info-box">✅ <b>{books_file.name}</b> uploaded successfully as <b>Books</b></div>""",
                    unsafe_allow_html=True)

# ── Run Reconciliation ────────────────────────────────────────────────────────
st.markdown("---")
run_btn = st.button("▶️ Run Reconciliation", type="primary", use_container_width=True)

if run_btn:
    if not gst_file or not books_file:
        st.error("❌ Please upload both files before running reconciliation.")
        st.stop()

    with st.spinner("Reading files and running reconciliation..."):
        try:
            # Read files
            df_gst,  gst_type  = read_gst_file(gst_path)
            df_books, _        = read_gst_file(books_path, "BOOKS")

            st.success(
                f"✅ {gst_type}: **{len(df_gst):,}** records  |  "
                f"Books: **{len(df_books):,}** records"
            )

            # Reconcile
            result_df = reconcile(df_gst, df_books, tolerance, date_window)
            summary   = build_summary(result_df, df_gst, df_books)
            source_label = df_gst["Source"].iloc[0]

        except Exception as e:
            st.error(f"❌ Error during reconciliation: {str(e)}")
            st.stop()

    # ── Summary Metrics ───────────────────────────────────────────────────────
    st.markdown("### 📊 Reconciliation Summary")

    mc = st.columns(5)
    status_counts = result_df["Status"].value_counts().to_dict()
    only_gst_key  = f"ONLY IN {source_label.upper()}"

    with mc[0]:
        st.metric("✅ Matched",          status_counts.get("MATCHED", 0))
    with mc[1]:
        st.metric("🔵 Probable Match",   status_counts.get("PROBABLE MATCH", 0))
    with mc[2]:
        st.metric("⚠️ Vendor Mismatch",  status_counts.get("VENDOR MISMATCH", 0))
    with mc[3]:
        st.metric(f"🟡 Only in {source_label}", status_counts.get(only_gst_key, 0))
    with mc[4]:
        st.metric("🟣 Only in Books",    status_counts.get("ONLY IN BOOKS", 0))

    # Financial summary
    st.markdown("### 💰 Financial Summary")
    fc = st.columns(3)
    with fc[0]:
        st.metric(f"Total ITC — {source_label}",
                  f"Rs {df_gst['Total_Tax'].sum():,.2f}")
    with fc[1]:
        st.metric("Total ITC — Books",
                  f"Rs {df_books['Total_Tax'].sum():,.2f}")
    with fc[2]:
        diff = df_books["Total_Tax"].sum() - df_gst["Total_Tax"].sum()
        st.metric("Net Difference",
                  f"Rs {abs(diff):,.2f}",
                  delta=f"{'Books > 2A' if diff>0 else '2A > Books'}",
                  delta_color="inverse")

    # ── Detailed Tables ───────────────────────────────────────────────────────
    st.markdown("### 📋 Detailed Results")
    tabs = st.tabs([
        "All Records", "✅ Matched", "🔵 Probable",
        "⚠️ Mismatch", f"🟡 Only {source_label}", "🟣 Only Books",
        "📈 Vendor Summary"
    ])

    display_cols = [
        "GSTIN","Supplier_Name","Invoice_No","Date_GST","Date_Books",
        "Voucher_No","GST_Total_Tax","Books_Total_Tax","Difference","Status"
    ]

    def fmt_df(df):
        cols = [c for c in display_cols if c in df.columns]
        return df[cols].reset_index(drop=True)

    with tabs[0]:
        st.dataframe(fmt_df(result_df), use_container_width=True, height=400)
    with tabs[1]:
        st.dataframe(fmt_df(result_df[result_df["Status"]=="MATCHED"]),
                     use_container_width=True, height=400)
    with tabs[2]:
        st.dataframe(fmt_df(result_df[result_df["Status"]=="PROBABLE MATCH"]),
                     use_container_width=True, height=400)
    with tabs[3]:
        st.dataframe(fmt_df(result_df[result_df["Status"]=="VENDOR MISMATCH"]),
                     use_container_width=True, height=400)
    with tabs[4]:
        st.dataframe(fmt_df(result_df[result_df["Status"]==only_gst_key]),
                     use_container_width=True, height=400)
    with tabs[5]:
        st.dataframe(fmt_df(result_df[result_df["Status"]=="ONLY IN BOOKS"]),
                     use_container_width=True, height=400)
    with tabs[6]:
        vendor_df = build_vendor_summary(result_df, source_label)
        st.dataframe(vendor_df, use_container_width=True, height=400)

    # ── Download Report ───────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### ⬇️ Download Excel Report")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
        out_path = tmp_out.name

    try:
        write_report(result_df, summary, df_gst, df_books, out_path, period, source_label)
        with open(out_path, "rb") as f:
            report_bytes = f.read()
        filename = f"GST_Recon_{source_label.replace('-','_')}_{period.replace(' ','_')}.xlsx"
        st.download_button(
            label="📥 Download Full Reconciliation Report (Excel)",
            data=report_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown(f"""<div class="info-box">
        ✅ Report contains 8 sheets: Summary | Recon Detail | Matched | Probable Match |
        Vendor Mismatch | Only in {source_label} | Only in Books | Vendor Summary
        </div>""", unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error generating report: {str(e)}")
    finally:
        try: os.unlink(out_path)
        except: pass

# ── Footer ─────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("""
<div style="text-align:center; color:#94A3B8; font-size:0.82rem; padding:12px 0;">
    GST Reconciliation Tool &nbsp;|&nbsp; Rule 36(4) CGST Rules &nbsp;|&nbsp;
    Built for Chartered Accountant Practice &nbsp;|&nbsp;
    <b>Data is not stored — fully secure</b>
</div>
""", unsafe_allow_html=True)
