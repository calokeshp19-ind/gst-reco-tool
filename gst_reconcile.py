#!/usr/bin/env python3
"""
========================================================================
GST Reconciliation Tool – Command Line Interface
========================================================================
Usage:
    python gst_reconcile.py \
        --gstr2b   path/to/gstr2b.xlsx \
        --books    path/to/purchase_register.xlsx \
        --output   GST_Recon_Report.xlsx \
        --tolerance 10 \
        --period   "Apr-2024 to Mar-2025"

========================================================================
"""

import argparse
import sys
import os

def run(gstr2b_path: str, books_path: str, output_path: str,
        tolerance: float = 10.0, period: str = ""):

    from gst_recon_engine import (
        read_file, detect_column_map, normalize_dataframe,
        GSTReconciler, build_summary,
    )
    from gst_recon_report import export_reconciliation_report

    print(f"\n{'='*60}")
    print("   GST Reconciliation Tool  |  CA Advisory")
    print(f"{'='*60}")
    print(f"  GSTR-2B File : {os.path.basename(gstr2b_path)}")
    print(f"  Books File   : {os.path.basename(books_path)}")
    print(f"  Tolerance    : ₹{tolerance:.2f}")
    print(f"  Period       : {period or 'Not specified'}")
    print(f"{'='*60}\n")

    # ── Read files ──
    print("[1/5] Reading input files...")
    raw_gstr2b = read_file(gstr2b_path)
    raw_books  = read_file(books_path)
    print(f"      GSTR-2B: {len(raw_gstr2b)} rows | {raw_gstr2b.shape[1]} columns")
    print(f"      Books  : {len(raw_books)} rows | {raw_books.shape[1]} columns")

    # ── Detect & map columns ──
    print("\n[2/5] Auto-detecting column mapping...")
    g_col_map = detect_column_map(raw_gstr2b)
    b_col_map = detect_column_map(raw_books)

    print("      GSTR-2B column map:")
    for k, v in g_col_map.items():
        print(f"        {k:22s} → {v}")
    print("      Books column map:")
    for k, v in b_col_map.items():
        print(f"        {k:22s} → {v}")

    # ── Normalize ──
    print("\n[3/5] Normalizing data...")
    gstr2b_norm = normalize_dataframe(raw_gstr2b, g_col_map, "GSTR2B")
    books_norm  = normalize_dataframe(raw_books,  b_col_map, "BOOKS")

    # Remove rows with empty GSTIN or Invoice_Number
    gstr2b_norm = gstr2b_norm[gstr2b_norm["Invoice_Number_Clean"].str.strip() != ""].reset_index(drop=True)
    books_norm  = books_norm[books_norm["Invoice_Number_Clean"].str.strip() != ""].reset_index(drop=True)
    print(f"      GSTR-2B valid rows : {len(gstr2b_norm)}")
    print(f"      Books valid rows   : {len(books_norm)}")

    # ── Reconcile ──
    print("\n[4/5] Running reconciliation engine...")
    recon = GSTReconciler(gstr2b_norm, books_norm, tolerance=tolerance)
    result_df = recon.reconcile()

    # Print status breakdown
    status_counts = result_df["Status"].value_counts()
    print("\n      ── Reconciliation Results ──")
    for status, count in status_counts.items():
        print(f"        {status:25s}: {count:5d} invoices")

    # ── Build summary ──
    summary_data = build_summary(result_df, gstr2b_norm, books_norm)
    overall = summary_data["overall"]
    print(f"\n      Total ITC as per GSTR-2B : ₹{overall['Total as per GSTR-2B (₹)']:>15,.2f}")
    print(f"      Total ITC as per Books   : ₹{overall['Total as per Books (₹)']:>15,.2f}")
    print(f"      Net Difference           : ₹{overall['Net Difference (₹)']:>15,.2f}")

    # ── Export ──
    print(f"\n[5/5] Generating Excel report → {output_path}")
    export_reconciliation_report(
        result_df, summary_data, gstr2b_norm, books_norm,
        output_path=output_path,
        period=period,
        tolerance=tolerance,
    )

    print(f"\n✅  Report saved: {output_path}")
    print(f"{'='*60}\n")
    return result_df, summary_data


# ─────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="GST Reconciliation Tool – GSTR-2B vs Purchase Register"
    )
    parser.add_argument("--gstr2b",    required=True, help="Path to GSTR-2B Excel file")
    parser.add_argument("--books",     required=True, help="Path to Purchase Register (Books) Excel file")
    parser.add_argument("--output",    default="GST_Reconciliation_Report.xlsx", help="Output Excel file path")
    parser.add_argument("--tolerance", type=float, default=10.0, help="Matching tolerance in ₹ (default: 10)")
    parser.add_argument("--period",    default="", help="Period label e.g. 'Apr-2024 to Mar-2025'")

    args = parser.parse_args()

    if not os.path.exists(args.gstr2b):
        print(f"❌  GSTR-2B file not found: {args.gstr2b}")
        sys.exit(1)
    if not os.path.exists(args.books):
        print(f"❌  Books file not found: {args.books}")
        sys.exit(1)

    run(
        gstr2b_path=args.gstr2b,
        books_path=args.books,
        output_path=args.output,
        tolerance=args.tolerance,
        period=args.period,
    )
