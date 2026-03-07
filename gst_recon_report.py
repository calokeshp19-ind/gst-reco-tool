"""
========================================================================
GST Reconciliation Report Generator
Produces a multi-sheet, colour-coded Excel workbook.
========================================================================
Sheets:
  1. Summary              – Overall + Category-wise totals
  2. Recon Detail         – Invoice-level reconciliation (all invoices)
  3. Only in GSTR-2B      – ITC in portal but not in books
  4. Only in Books        – ITC in books but not in portal
  5. Mismatches           – Matched but with amount difference > tolerance
  6. Vendor Summary       – GSTIN-wise ITC comparison
  7. GSTR-2B Raw          – Original uploaded data
  8. Books Raw            – Original uploaded data
========================================================================
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────
# COLOUR PALETTE
# ─────────────────────────────────────────────────────────────────────
C = {
    "header_dark":   "1E3A5F",   # Navy  – primary header
    "header_mid":    "2563EB",   # Blue  – sub-header
    "header_green":  "065F46",   # Dark green
    "header_red":    "7F1D1D",   # Dark red
    "header_amber":  "78350F",   # Amber
    "header_purple": "4C1D95",   # Purple

    "matched":       "D1FAE5",   # Light green
    "mismatch":      "FEE2E2",   # Light red
    "only2b":        "FEF3C7",   # Light amber
    "only_books":    "EDE9FE",   # Light purple
    "probable":      "DBEAFE",   # Light blue
    "cn":            "FCE7F3",   # Light pink
    "amd":           "FFF7ED",   # Light orange

    "alt_row":       "F8FAFC",   # Alternate row
    "white":         "FFFFFF",
    "total_row":     "1E3A5F",   # Navy for total rows
    "total_text":    "FFFFFF",
}

STATUS_FILL = {
    "MATCHED":           C["matched"],
    "CN MATCHED":        C["cn"],
    "AMENDMENT MATCH":   C["amd"],
    "PROBABLE MATCH":    C["probable"],
    "MISMATCH":          C["mismatch"],
    "ONLY IN GSTR-2B":   C["only2b"],
    "ONLY IN BOOKS":     C["only_books"],
}

FONT_BODY   = "Arial"
FONT_HEAD   = "Arial"


# ─────────────────────────────────────────────────────────────────────
# STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_color, end_color=hex_color)

def _font(bold=False, color="000000", size=10, name=FONT_BODY) -> Font:
    return Font(name=name, bold=bold, color=color, size=size)

def _align(h="left", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _border_thin() -> Border:
    s = Side(style="thin", color="CBD5E1")
    return Border(left=s, right=s, top=s, bottom=s)

def _border_medium() -> Border:
    s = Side(style="medium", color="94A3B8")
    return Border(left=s, right=s, top=s, bottom=s)

def _apply_header_row(ws, row_num: int, values: list, bg: str, font_color="FFFFFF",
                       font_size=10, bold=True):
    for col, val in enumerate(values, start=1):
        cell = ws.cell(row=row_num, column=col, value=val)
        cell.fill   = _fill(bg)
        cell.font   = _font(bold=bold, color=font_color, size=font_size, name=FONT_HEAD)
        cell.alignment = _align(h="center", wrap=True)
        cell.border = _border_thin()

def _num_fmt(ws, row, col, fmt="#,##0.00"):
    ws.cell(row=row, column=col).number_format = fmt

def _set_col_widths(ws, widths: list):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def _freeze(ws, cell="B2"):
    ws.freeze_panes = cell


# ─────────────────────────────────────────────────────────────────────
# TITLE BLOCK
# ─────────────────────────────────────────────────────────────────────

def _write_title(ws, title: str, subtitle: str, period: str = ""):
    ws.merge_cells("A1:M1")
    c = ws["A1"]
    c.value = title
    c.fill  = _fill(C["header_dark"])
    c.font  = _font(bold=True, color="FFFFFF", size=14, name=FONT_HEAD)
    c.alignment = _align(h="center")

    ws.merge_cells("A2:M2")
    c = ws["A2"]
    c.value = subtitle + (f"  |  Period: {period}" if period else "")
    c.fill  = _fill(C["header_mid"])
    c.font  = _font(bold=False, color="FFFFFF", size=10)
    c.alignment = _align(h="center")

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18


# ─────────────────────────────────────────────────────────────────────
# SHEET 1 – SUMMARY
# ─────────────────────────────────────────────────────────────────────

def write_summary_sheet(ws, summary_data: dict, tolerance: float, period: str = ""):
    ws.title = "Summary"
    _write_title(ws, "GST Reconciliation – Summary", "GSTR-2B vs Books of Accounts", period)

    # ── Overall KPIs ──
    ws["A4"] = "OVERALL RECONCILIATION"
    ws["A4"].font = _font(bold=True, color=C["header_dark"], size=11)
    ws["A4"].fill = _fill("EFF6FF")
    ws.merge_cells("A4:D4")
    ws["A4"].alignment = _align(h="left")

    headers = ["Particulars", "Amount (₹)"]
    _apply_header_row(ws, 5, headers + ["", ""], C["header_dark"])

    overall = summary_data["overall"]
    kpi_rows = [
        ("Total ITC as per GSTR-2B",     overall["Total as per GSTR-2B (₹)"],    C["matched"]),
        ("Total ITC as per Books",        overall["Total as per Books (₹)"],      C["probable"]),
        ("Net Difference (Books – 2B)",   overall["Net Difference (₹)"],          C["mismatch"] if overall["Net Difference (₹)"] != 0 else C["matched"]),
        (None, None, None),
        ("Matched Invoices (count)",      overall["Matched"],                     C["matched"]),
        ("CN Matched (count)",            overall["CN Matched"],                  C["cn"]),
        ("Amendment Match (count)",       overall["Amendment Match"],             C["amd"]),
        ("Probable Match (count)",        overall["Probable Match"],              C["probable"]),
        ("Mismatch (count)",              overall["Mismatch"],                    C["mismatch"]),
        (None, None, None),
        ("Only in GSTR-2B (count)",       overall["Only in GSTR-2B (count)"],    C["only2b"]),
        ("Only in GSTR-2B (₹)",          overall["Only in GSTR-2B (₹)"],        C["only2b"]),
        ("Only in Books (count)",         overall["Only in Books (count)"],       C["only_books"]),
        ("Only in Books (₹)",            overall["Only in Books (₹)"],           C["only_books"]),
    ]

    r = 6
    for label, val, fill_hex in kpi_rows:
        if label is None:
            r += 1
            continue
        c1 = ws.cell(row=r, column=1, value=label)
        c2 = ws.cell(row=r, column=2, value=val)
        for c in [c1, c2]:
            c.fill = _fill(fill_hex)
            c.font = _font(size=10)
            c.border = _border_thin()
            c.alignment = _align(h="left")
        if isinstance(val, float):
            c2.number_format = "#,##0.00"
            c2.alignment = _align(h="right")
        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 20
        r += 1

    # Tolerance note
    r += 1
    ws.cell(row=r, column=1, value=f"Note: Matching tolerance set at ₹{tolerance:.2f}. Refer Rule 36(4) CGST Rules & Circular 183/15/2022-GST")
    ws.cell(row=r, column=1).font = _font(size=9, color="64748B")
    ws.merge_cells(f"A{r}:G{r}")

    r += 1
    ws.cell(row=r, column=1, value=f"Generated on: {datetime.now().strftime('%d-%b-%Y %H:%M')}")
    ws.cell(row=r, column=1).font = _font(size=9, color="64748B")

    # ── Category Table ──
    r += 2
    ws.cell(row=r, column=1, value="CATEGORY-WISE BREAKUP")
    ws.cell(row=r, column=1).font = _font(bold=True, color=C["header_dark"], size=11)
    ws.cell(row=r, column=1).fill = _fill("EFF6FF")
    ws.merge_cells(f"A{r}:J{r}")
    r += 1

    cat_df = summary_data["category"]
    cat_headers = list(cat_df.columns)
    _apply_header_row(ws, r, cat_headers, C["header_mid"])
    r += 1

    for _, row_data in cat_df.iterrows():
        for col_i, val in enumerate(row_data.values, start=1):
            c = ws.cell(row=r, column=col_i, value=val)
            c.border = _border_thin()
            c.font   = _font(size=10)
            c.alignment = _align(h="right" if col_i > 1 else "left")
            if isinstance(val, float):
                c.number_format = "#,##0.00"
        r += 1

    _freeze(ws, "A3")


# ─────────────────────────────────────────────────────────────────────
# SHEET 2 – RECON DETAIL (Invoice Level)
# ─────────────────────────────────────────────────────────────────────

DETAIL_COLS = [
    ("GSTIN",              20, "left"),
    ("Supplier_Name",      28, "left"),
    ("Invoice_Number",     18, "left"),
    ("Invoice_Date",       14, "center"),
    ("Doc_Category",        8, "center"),
    ("GSTR2B_Taxable",     16, "right"),
    ("GSTR2B_IGST",        12, "right"),
    ("GSTR2B_CGST",        12, "right"),
    ("GSTR2B_SGST",        12, "right"),
    ("GSTR2B_Total_Tax",   16, "right"),
    ("Books_Taxable",      16, "right"),
    ("Books_IGST",         12, "right"),
    ("Books_CGST",         12, "right"),
    ("Books_SGST",         12, "right"),
    ("Books_Total_Tax",    16, "right"),
    ("Difference_Taxable", 16, "right"),
    ("Difference_Tax",     16, "right"),
    ("Status",             22, "center"),
]

DISPLAY_HEADERS = [
    "GSTIN", "Supplier Name", "Invoice No", "Invoice Date", "Doc Type",
    "2B Taxable (₹)", "2B IGST (₹)", "2B CGST (₹)", "2B SGST (₹)", "2B Total Tax (₹)",
    "Books Taxable (₹)", "Books IGST (₹)", "Books CGST (₹)", "Books SGST (₹)", "Books Total Tax (₹)",
    "Diff Taxable (₹)", "Diff Tax (₹)", "Status",
]

NUM_COLS = {5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16}  # 0-indexed


def _write_detail_sheet(ws, df: pd.DataFrame, sheet_title: str,
                         subtitle: str, header_color: str, period: str = ""):
    ws.title = sheet_title
    _write_title(ws, sheet_title, subtitle, period)

    # Column widths
    for i, (_, w, _) in enumerate(DETAIL_COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Header row
    _apply_header_row(ws, 4, DISPLAY_HEADERS, header_color)
    ws.row_dimensions[4].height = 30

    if df.empty:
        ws.cell(row=5, column=1, value="No records.")
        return

    col_names = [c[0] for c in DETAIL_COLS]

    for r_idx, (_, row_data) in enumerate(df.iterrows(), start=5):
        fill_hex = STATUS_FILL.get(str(row_data.get("Status", "")), C["white"])
        alt_fill = C["alt_row"] if r_idx % 2 == 0 else C["white"]

        for c_idx, col in enumerate(col_names, start=1):
            val = row_data.get(col, None)
            # Format dates
            if col == "Invoice_Date" and pd.notna(val):
                try:
                    val = pd.to_datetime(val).strftime("%d-%b-%Y")
                except:
                    pass
            # NaN → blank
            if isinstance(val, float) and np.isnan(val):
                val = None

            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border = _border_thin()
            cell.font   = _font(size=9)
            cell.fill   = _fill(fill_hex if fill_hex else alt_fill)

            _, _, align = DETAIL_COLS[c_idx - 1]
            cell.alignment = _align(h=align)

            if c_idx - 1 in NUM_COLS and val is not None:
                cell.number_format = "#,##0.00"

    # Auto-filter
    ws.auto_filter.ref = f"A4:{get_column_letter(len(DETAIL_COLS))}4"
    _freeze(ws, "A5")


# ─────────────────────────────────────────────────────────────────────
# SHEET 6 – VENDOR SUMMARY
# ─────────────────────────────────────────────────────────────────────

def write_vendor_summary(ws, vendor_df: pd.DataFrame, period: str = ""):
    ws.title = "Vendor Summary"
    _write_title(ws, "Vendor-wise ITC Summary", "GSTR-2B vs Books of Accounts", period)

    headers = ["GSTIN", "Vendor Name", "Books ITC (₹)", "GSTR-2B ITC (₹)", "Difference (₹)"]
    _apply_header_row(ws, 4, headers, C["header_dark"])
    widths = [22, 30, 18, 18, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 5
    for _, row_data in vendor_df.iterrows():
        diff = row_data.get("Difference", 0) or 0
        fill = C["matched"] if abs(diff) <= 10 else (C["mismatch"] if diff < 0 else C["only_books"])

        vals = [
            row_data.get("GSTIN", ""),
            row_data.get("Supplier_Name", ""),
            row_data.get("Books_ITC", 0),
            row_data.get("GSTR2B_ITC", 0),
            diff,
        ]
        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.border = _border_thin()
            cell.font   = _font(size=10)
            cell.fill   = _fill(fill if c_idx == 5 else (C["alt_row"] if r % 2 == 0 else C["white"]))
            cell.alignment = _align(h="right" if c_idx > 2 else "left")
            if c_idx > 2:
                cell.number_format = "#,##0.00"
        r += 1

    # Total row
    if r > 5:
        total_row = r
        ws.cell(total_row, 1, "TOTAL").font = _font(bold=True, color="FFFFFF", size=10)
        ws.cell(total_row, 1).fill = _fill(C["total_row"])
        ws.cell(total_row, 1).alignment = _align(h="center")
        for c_idx in range(2, 6):
            ws.cell(total_row, c_idx).fill = _fill(C["total_row"])
        ws.cell(total_row, 3).value = f"=SUM(C5:C{total_row-1})"
        ws.cell(total_row, 4).value = f"=SUM(D5:D{total_row-1})"
        ws.cell(total_row, 5).value = f"=SUM(E5:E{total_row-1})"
        for c_idx in [3, 4, 5]:
            ws.cell(total_row, c_idx).number_format = "#,##0.00"
            ws.cell(total_row, c_idx).font = _font(bold=True, color="FFFFFF", size=10)
            ws.cell(total_row, c_idx).alignment = _align(h="right")

    ws.auto_filter.ref = f"A4:E4"
    _freeze(ws, "A5")


# ─────────────────────────────────────────────────────────────────────
# RAW DATA SHEETS
# ─────────────────────────────────────────────────────────────────────

def write_raw_sheet(ws, df: pd.DataFrame, title: str):
    ws.title = title
    if df.empty:
        ws["A1"] = "No data."
        return
    headers = list(df.columns)
    _apply_header_row(ws, 1, headers, C["header_mid"])
    for i, col in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 18

    for r_idx, (_, row_data) in enumerate(df.iterrows(), start=2):
        fill = C["alt_row"] if r_idx % 2 == 0 else C["white"]
        for c_idx, val in enumerate(row_data.values, start=1):
            if isinstance(val, float) and np.isnan(val):
                val = None
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font   = _font(size=9)
            cell.fill   = _fill(fill)
            cell.border = _border_thin()
            cell.alignment = _align()

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"
    _freeze(ws, "A2")


# ─────────────────────────────────────────────────────────────────────
# MASTER EXPORT FUNCTION
# ─────────────────────────────────────────────────────────────────────

def export_reconciliation_report(
    result_df: pd.DataFrame,
    summary_data: dict,
    gstr2b_df: pd.DataFrame,
    books_df: pd.DataFrame,
    output_path: str,
    period: str = "",
    tolerance: float = 10.0,
) -> str:
    """
    Generate full reconciliation Excel workbook.
    Returns output_path.
    """
    wb = Workbook()

    # ── Sheet 1: Summary ──
    ws1 = wb.active
    write_summary_sheet(ws1, summary_data, tolerance, period)

    # ── Sheet 2: Recon Detail ──
    ws2 = wb.create_sheet("Recon Detail")
    _write_detail_sheet(
        ws2, result_df,
        "Recon Detail", "Invoice-Level Reconciliation – All Records",
        C["header_dark"], period
    )

    # ── Sheet 3: Only in GSTR-2B ──
    df_2b = result_df[result_df["Status"] == "ONLY IN GSTR-2B"].copy()
    ws3 = wb.create_sheet("Only in GSTR-2B")
    _write_detail_sheet(
        ws3, df_2b,
        "Only in GSTR-2B", "ITC Available in Portal – NOT in Books",
        C["header_amber"], period
    )

    # ── Sheet 4: Only in Books ──
    df_bk = result_df[result_df["Status"] == "ONLY IN BOOKS"].copy()
    ws4 = wb.create_sheet("Only in Books")
    _write_detail_sheet(
        ws4, df_bk,
        "Only in Books", "ITC Claimed in Books – NOT reflected in GSTR-2B",
        C["header_purple"], period
    )

    # ── Sheet 5: Mismatches ──
    df_mm = result_df[result_df["Status"] == "MISMATCH"].copy()
    ws5 = wb.create_sheet("Mismatches")
    _write_detail_sheet(
        ws5, df_mm,
        "Mismatches", "Matched by Invoice No but Amount Difference > Tolerance",
        C["header_red"], period
    )

    # ── Sheet 6: Vendor Summary ──
    from gst_recon_engine import build_vendor_summary
    vendor_df = build_vendor_summary(result_df)
    ws6 = wb.create_sheet("Vendor Summary")
    write_vendor_summary(ws6, vendor_df, period)

    # ── Sheet 7: GSTR-2B Raw ──
    ws7 = wb.create_sheet("GSTR-2B Raw")
    write_raw_sheet(ws7, gstr2b_df, "GSTR-2B Raw")

    # ── Sheet 8: Books Raw ──
    ws8 = wb.create_sheet("Books Raw")
    write_raw_sheet(ws8, books_df, "Books Raw")

    wb.save(output_path)
    return output_path
