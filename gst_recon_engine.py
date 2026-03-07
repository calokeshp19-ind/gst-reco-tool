"""
========================================================================
GST Reconciliation Engine  –  GSTR-2B vs Purchase Register (Books)
========================================================================
Author : CA Advisory Tool
Version: 2.0
Purpose: Reconcile ITC as per GSTR-2B against Books of Accounts.
         Handles B2B invoices, Credit Notes, Debit Notes & Amendments.

Matching Logic (in order of priority):
  1. PRIMARY   – GSTIN + Invoice_Number  (taxable diff ≤ tolerance)
  2. CN MATCH  – GSTIN + CN No + Taxable Value (for Credit Notes)
  3. AMENDMENT – GSTIN + Original Invoice No + Taxable Value
  4. PROBABLE  – GSTIN + Taxable Value + Invoice Date within ±5 days
  5. ONLY IN GSTR-2B / ONLY IN BOOKS (unmatched residuals)

Ref:
  • CGST Act 2017 – Section 16 (ITC Eligibility)
  • Rule 36(4) CGST Rules – ITC claim basis GSTR-2B
  • GST Circular No. 183/15/2022-GST (GSTR-2B reconciliation)
========================================================================
"""

import re
import warnings
import numpy as np
import pandas as pd
from datetime import timedelta

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────
# 1.  COLUMN ALIAS MAPS
# ─────────────────────────────────────────────────────────────────────

COLUMN_ALIASES = {
    "GSTIN": [
        "supplier gstin", "gstin of supplier", "vendor gstin",
        "gstin", "party gstin", "gst no", "gst number",
    ],
    "Invoice_Number": [
        "invoice number", "invoice no", "inv no", "bill no",
        "document number", "document no", "doc no", "voucher no",
        "bill number", "reference no", "ref no",
    ],
    "Invoice_Date": [
        "invoice date", "inv date", "bill date", "document date",
        "doc date", "date", "voucher date",
    ],
    "Taxable_Value": [
        "taxable value", "taxable amount", "assessable value",
        "taxable val", "basic value", "taxable", "base amount",
    ],
    "IGST": [
        "igst amount", "igst", "integrated tax", "integrated gst",
    ],
    "CGST": [
        "cgst amount", "cgst", "central tax", "central gst",
    ],
    "SGST": [
        "sgst amount", "sgst", "sgst/utgst", "state/ut tax",
        "state tax", "utgst",
    ],
    "Cess": [
        "cess amount", "cess",
    ],
    "Supplier_Name": [
        "trade/legal name", "supplier name", "vendor name",
        "party name", "name of supplier", "supplier", "vendor",
    ],
    "Document_Type": [
        "document type", "doc type", "type", "invoice type",
        "transaction type", "note type",
    ],
    "Invoice_Value": [
        "invoice value", "total invoice value", "bill amount",
        "total amount", "gross amount",
    ],
}

# ─────────────────────────────────────────────────────────────────────
# 2.  COLUMN DETECTION & NORMALIZATION
# ─────────────────────────────────────────────────────────────────────

def _norm_header(h: str) -> str:
    """Lowercase, strip special chars for fuzzy matching."""
    return re.sub(r"[^a-z0-9 /]", " ", str(h).lower()).strip()


def detect_column_map(df: pd.DataFrame) -> dict:
    """
    Auto-detect which DataFrame columns correspond to standard fields.
    Returns dict  {standard_field: actual_column_name}.
    """
    norm_cols = {_norm_header(c): c for c in df.columns}
    mapping = {}
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            norm_alias = _norm_header(alias)
            if norm_alias in norm_cols:
                mapping[field] = norm_cols[norm_alias]
                break
            # partial match fallback
            for norm_c, actual_c in norm_cols.items():
                if norm_alias in norm_c or norm_c in norm_alias:
                    if field not in mapping:
                        mapping[field] = actual_c
    return mapping


def normalize_dataframe(df: pd.DataFrame, col_map: dict, source: str) -> pd.DataFrame:
    """
    Rename detected columns to standard names, compute Total_Tax,
    and classify Document_Type.
    source: 'GSTR2B' | 'BOOKS'
    """
    df = df.copy()

    # Rename to standard names
    rename = {v: k for k, v in col_map.items() if k in COLUMN_ALIASES}
    df = df.rename(columns=rename)

    # Ensure all standard numeric columns exist
    for col in ["Taxable_Value", "IGST", "CGST", "SGST", "Cess", "Invoice_Value"]:
        if col not in df.columns:
            df[col] = 0.0

    # Numeric coercion
    for col in ["Taxable_Value", "IGST", "CGST", "SGST", "Cess", "Invoice_Value"]:
        df[col] = pd.to_numeric(
            df[col].astype(str).str.replace(",", "").str.strip(),
            errors="coerce"
        ).fillna(0.0)

    # Compute Total_Tax
    df["Total_Tax"] = df["IGST"] + df["CGST"] + df["SGST"] + df["Cess"]

    # Classify Document_Type
    if "Document_Type" not in df.columns:
        df["Document_Type"] = "INV"
    df["Document_Type"] = df["Document_Type"].fillna("INV").astype(str).str.strip()
    df["Doc_Category"] = df["Document_Type"].apply(_classify_doc_type)

    # Parse Invoice_Date
    if "Invoice_Date" in df.columns:
        df["Invoice_Date"] = pd.to_datetime(df["Invoice_Date"], errors="coerce", dayfirst=True)
    else:
        df["Invoice_Date"] = pd.NaT

    # Ensure string cols
    for col in ["GSTIN", "Invoice_Number", "Supplier_Name"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).str.strip()

    # Standardize invoice numbers
    df["Invoice_Number_Clean"] = df["Invoice_Number"].apply(clean_invoice_number)

    df["_source"] = source
    df = df.reset_index(drop=True)
    return df


def _classify_doc_type(val: str) -> str:
    """Map raw document type string to category: INV / CN / DN / AMD."""
    v = str(val).lower()
    if any(x in v for x in ["credit", "cr note", "cn"]):
        return "CN"
    if any(x in v for x in ["debit", "dn"]):
        return "DN"
    if "amend" in v:
        return "AMD"
    return "INV"


# ─────────────────────────────────────────────────────────────────────
# 3.  INVOICE NUMBER CLEANING
# ─────────────────────────────────────────────────────────────────────

def clean_invoice_number(inv: str) -> str:
    """
    Standardize invoice number:
      - Remove hyphens, slashes, spaces
      - Uppercase
    Example: INV-00123/24 A → INV0012324A
    """
    cleaned = re.sub(r"[-/\s]", "", str(inv).upper())
    return cleaned.strip()


# ─────────────────────────────────────────────────────────────────────
# 4.  MAIN RECONCILIATION ENGINE
# ─────────────────────────────────────────────────────────────────────

class GSTReconciler:
    """
    Core reconciliation engine.
    Usage:
        recon = GSTReconciler(gstr2b_df, books_df, tolerance=10)
        result = recon.reconcile()
    """

    def __init__(
        self,
        gstr2b_df: pd.DataFrame,
        books_df: pd.DataFrame,
        tolerance: float = 10.0,
        probable_date_window: int = 5,
    ):
        self.gstr2b = gstr2b_df.copy()
        self.books  = books_df.copy()
        self.tol    = tolerance
        self.date_window = probable_date_window

        # Track which rows have been matched (by index)
        self._gstr2b_matched = set()
        self._books_matched  = set()

        self.result_rows = []   # list of dicts → final reconciliation table

    # ── Public entry point ──────────────────────────────────────────
    def reconcile(self) -> pd.DataFrame:
        self._step1_primary_match()
        self._step2_cn_match()
        self._step3_amendment_match()
        self._step4_probable_match()
        self._step5_residuals()
        return self._build_result_df()

    # ── Step 1 : Primary Match (GSTIN + Invoice_Number) ─────────────
    def _step1_primary_match(self):
        g = self.gstr2b.copy()
        b = self.books.copy()

        merged = pd.merge(
            g.reset_index().rename(columns={"index": "g_idx"}),
            b.reset_index().rename(columns={"index": "b_idx"}),
            left_on=["GSTIN", "Invoice_Number_Clean"],
            right_on=["GSTIN", "Invoice_Number_Clean"],
            suffixes=("_g", "_b"),
        )

        for _, row in merged.iterrows():
            if row["g_idx"] in self._gstr2b_matched:
                continue
            if row["b_idx"] in self._books_matched:
                continue

            diff = abs(row["Taxable_Value_b"] - row["Taxable_Value_g"])
            status = "MATCHED" if diff <= self.tol else "MISMATCH"

            self._record(row, status, "PRIMARY", diff)
            self._gstr2b_matched.add(row["g_idx"])
            self._books_matched.add(row["b_idx"])

    # ── Step 2 : Credit Note Match ──────────────────────────────────
    def _step2_cn_match(self):
        g_cn = self.gstr2b[
            (self.gstr2b["Doc_Category"] == "CN") &
            (~self.gstr2b.index.isin(self._gstr2b_matched))
        ].copy()
        b_cn = self.books[
            (self.books["Doc_Category"] == "CN") &
            (~self.books.index.isin(self._books_matched))
        ].copy()

        merged = pd.merge(
            g_cn.reset_index().rename(columns={"index": "g_idx"}),
            b_cn.reset_index().rename(columns={"index": "b_idx"}),
            on=["GSTIN", "Invoice_Number_Clean"],
            suffixes=("_g", "_b"),
        )

        for _, row in merged.iterrows():
            if row["g_idx"] in self._gstr2b_matched:
                continue
            if row["b_idx"] in self._books_matched:
                continue

            diff = abs(row["Taxable_Value_b"] - row["Taxable_Value_g"])
            if diff <= self.tol:
                self._record(row, "CN MATCHED", "CN", diff)
                self._gstr2b_matched.add(row["g_idx"])
                self._books_matched.add(row["b_idx"])

    # ── Step 3 : Amendment Match ────────────────────────────────────
    def _step3_amendment_match(self):
        g_amd = self.gstr2b[
            (self.gstr2b["Doc_Category"] == "AMD") &
            (~self.gstr2b.index.isin(self._gstr2b_matched))
        ].copy()
        b_amd = self.books[
            (~self.books.index.isin(self._books_matched))
        ].copy()

        merged = pd.merge(
            g_amd.reset_index().rename(columns={"index": "g_idx"}),
            b_amd.reset_index().rename(columns={"index": "b_idx"}),
            on=["GSTIN", "Invoice_Number_Clean"],
            suffixes=("_g", "_b"),
        )

        for _, row in merged.iterrows():
            if row["g_idx"] in self._gstr2b_matched:
                continue
            if row["b_idx"] in self._books_matched:
                continue

            diff = abs(row["Taxable_Value_b"] - row["Taxable_Value_g"])
            if diff <= self.tol:
                self._record(row, "AMENDMENT MATCH", "AMENDMENT", diff)
                self._gstr2b_matched.add(row["g_idx"])
                self._books_matched.add(row["b_idx"])

    # ── Step 4 : Probable Match (GSTIN + Taxable Value + Date ±5d) ──
    def _step4_probable_match(self):
        g_rem = self.gstr2b[~self.gstr2b.index.isin(self._gstr2b_matched)].copy()
        b_rem = self.books[~self.books.index.isin(self._books_matched)].copy()

        merged = pd.merge(
            g_rem.reset_index().rename(columns={"index": "g_idx"}),
            b_rem.reset_index().rename(columns={"index": "b_idx"}),
            on="GSTIN",
            suffixes=("_g", "_b"),
        )

        if merged.empty:
            return

        # Filter on taxable value tolerance
        merged = merged[
            abs(merged["Taxable_Value_b"] - merged["Taxable_Value_g"]) <= self.tol
        ]

        # Date proximity filter
        def _date_ok(row):
            dg = row.get("Invoice_Date_g")
            db = row.get("Invoice_Date_b")
            if pd.isnull(dg) or pd.isnull(db):
                return True  # can't rule out; give benefit of doubt
            return abs((dg - db).days) <= self.date_window

        if "Invoice_Date_g" in merged.columns and "Invoice_Date_b" in merged.columns:
            merged = merged[merged.apply(_date_ok, axis=1)]

        for _, row in merged.iterrows():
            if row["g_idx"] in self._gstr2b_matched:
                continue
            if row["b_idx"] in self._books_matched:
                continue

            diff = abs(row["Taxable_Value_b"] - row["Taxable_Value_g"])
            self._record(row, "PROBABLE MATCH", "PROBABLE", diff)
            self._gstr2b_matched.add(row["g_idx"])
            self._books_matched.add(row["b_idx"])

    # ── Step 5 : Residuals ──────────────────────────────────────────
    def _step5_residuals(self):
        # Only in GSTR-2B
        for idx, row in self.gstr2b[
            ~self.gstr2b.index.isin(self._gstr2b_matched)
        ].iterrows():
            self.result_rows.append({
                "GSTIN":              row.get("GSTIN", ""),
                "Supplier_Name":      row.get("Supplier_Name", ""),
                "Invoice_Number":     row.get("Invoice_Number", ""),
                "Invoice_Date":       row.get("Invoice_Date", pd.NaT),
                "Doc_Category":       row.get("Doc_Category", "INV"),
                "GSTR2B_Taxable":     row.get("Taxable_Value", 0),
                "GSTR2B_IGST":        row.get("IGST", 0),
                "GSTR2B_CGST":        row.get("CGST", 0),
                "GSTR2B_SGST":        row.get("SGST", 0),
                "GSTR2B_Total_Tax":   row.get("Total_Tax", 0),
                "Books_Taxable":      None,
                "Books_IGST":         None,
                "Books_CGST":         None,
                "Books_SGST":         None,
                "Books_Total_Tax":    None,
                "Difference_Taxable": None,
                "Difference_Tax":     None,
                "Status":             "ONLY IN GSTR-2B",
                "Match_Type":         "NONE",
            })

        # Only in Books
        for idx, row in self.books[
            ~self.books.index.isin(self._books_matched)
        ].iterrows():
            self.result_rows.append({
                "GSTIN":              row.get("GSTIN", ""),
                "Supplier_Name":      row.get("Supplier_Name", ""),
                "Invoice_Number":     row.get("Invoice_Number", ""),
                "Invoice_Date":       row.get("Invoice_Date", pd.NaT),
                "Doc_Category":       row.get("Doc_Category", "INV"),
                "GSTR2B_Taxable":     None,
                "GSTR2B_IGST":        None,
                "GSTR2B_CGST":        None,
                "GSTR2B_SGST":        None,
                "GSTR2B_Total_Tax":   None,
                "Books_Taxable":      row.get("Taxable_Value", 0),
                "Books_IGST":         row.get("IGST", 0),
                "Books_CGST":         row.get("CGST", 0),
                "Books_SGST":         row.get("SGST", 0),
                "Books_Total_Tax":    row.get("Total_Tax", 0),
                "Difference_Taxable": None,
                "Difference_Tax":     None,
                "Status":             "ONLY IN BOOKS",
                "Match_Type":         "NONE",
            })

    # ── Internal record builder ─────────────────────────────────────
    def _record(self, row: pd.Series, status: str, match_type: str, diff: float):
        """Append a matched row result."""

        def _g(col):
            for suffix in [f"_{col}_g", f"_{col}", f"_{col}_x"]:
                if col + "_g" in row.index: return row[col + "_g"]
            return row.get(col, None)

        def _b(col):
            if col + "_b" in row.index: return row[col + "_b"]
            return row.get(col, None)

        # Resolve with suffix
        def get(base, side):
            key = f"{base}_{side}"
            return row[key] if key in row.index else row.get(base, None)

        g_tax  = get("Taxable_Value", "g")
        b_tax  = get("Taxable_Value", "b")
        g_igst = get("IGST", "g")
        b_igst = get("IGST", "b")
        g_cgst = get("CGST", "g")
        b_cgst = get("CGST", "b")
        g_sgst = get("SGST", "g")
        b_sgst = get("SGST", "b")
        g_ttax = get("Total_Tax", "g")
        b_ttax = get("Total_Tax", "b")

        gstin   = row.get("GSTIN", "")
        inv_no  = row.get("Invoice_Number_Clean", row.get("Invoice_Number", ""))
        inv_date= get("Invoice_Date", "g") or get("Invoice_Date", "b")
        sup_name= get("Supplier_Name", "g") or get("Supplier_Name", "b") or ""
        doc_cat = get("Doc_Category", "g") or "INV"

        d_tax = (b_tax - g_tax) if (b_tax is not None and g_tax is not None) else None
        d_ttax= (b_ttax - g_ttax) if (b_ttax is not None and g_ttax is not None) else None

        self.result_rows.append({
            "GSTIN":              gstin,
            "Supplier_Name":      sup_name,
            "Invoice_Number":     inv_no,
            "Invoice_Date":       inv_date,
            "Doc_Category":       doc_cat,
            "GSTR2B_Taxable":     g_tax,
            "GSTR2B_IGST":        g_igst,
            "GSTR2B_CGST":        g_cgst,
            "GSTR2B_SGST":        g_sgst,
            "GSTR2B_Total_Tax":   g_ttax,
            "Books_Taxable":      b_tax,
            "Books_IGST":         b_igst,
            "Books_CGST":         b_cgst,
            "Books_SGST":         b_sgst,
            "Books_Total_Tax":    b_ttax,
            "Difference_Taxable": d_tax,
            "Difference_Tax":     d_ttax,
            "Status":             status,
            "Match_Type":         match_type,
        })

    # ── Build final DataFrame ───────────────────────────────────────
    def _build_result_df(self) -> pd.DataFrame:
        if not self.result_rows:
            return pd.DataFrame()
        df = pd.DataFrame(self.result_rows)
        # Clean up numeric nulls
        num_cols = [
            "GSTR2B_Taxable","GSTR2B_IGST","GSTR2B_CGST","GSTR2B_SGST","GSTR2B_Total_Tax",
            "Books_Taxable","Books_IGST","Books_CGST","Books_SGST","Books_Total_Tax",
            "Difference_Taxable","Difference_Tax",
        ]
        for c in num_cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        return df


# ─────────────────────────────────────────────────────────────────────
# 5.  SUMMARY BUILDER
# ─────────────────────────────────────────────────────────────────────

def build_summary(result_df: pd.DataFrame, gstr2b_df: pd.DataFrame, books_df: pd.DataFrame) -> dict:
    """
    Compute overall and category-wise summary figures.
    Returns dict of DataFrames.
    """
    total_2b    = gstr2b_df["Total_Tax"].sum()
    total_books = books_df["Total_Tax"].sum()

    status_counts = result_df["Status"].value_counts().to_dict()

    # Category-wise
    cat_rows = []
    for cat in ["INV", "CN", "DN", "AMD"]:
        sub = result_df[result_df["Doc_Category"] == cat]
        cat_rows.append({
            "Category":            cat,
            "Total as per GSTR-2B (₹)":  _safe_sum(sub, "GSTR2B_Total_Tax"),
            "Total as per Books (₹)":    _safe_sum(sub, "Books_Total_Tax"),
            "Matched":             len(sub[sub["Status"] == "MATCHED"]),
            "CN Matched":          len(sub[sub["Status"] == "CN MATCHED"]),
            "Amendment Match":     len(sub[sub["Status"] == "AMENDMENT MATCH"]),
            "Probable Match":      len(sub[sub["Status"] == "PROBABLE MATCH"]),
            "Mismatch":            len(sub[sub["Status"] == "MISMATCH"]),
            "Only in GSTR-2B":     len(sub[sub["Status"] == "ONLY IN GSTR-2B"]),
            "Only in Books":       len(sub[sub["Status"] == "ONLY IN BOOKS"]),
        })

    category_df = pd.DataFrame(cat_rows)

    overall = {
        "Total as per GSTR-2B (₹)":     round(total_2b, 2),
        "Total as per Books (₹)":        round(total_books, 2),
        "Net Difference (₹)":            round(total_books - total_2b, 2),
        "Matched":                       status_counts.get("MATCHED", 0),
        "CN Matched":                    status_counts.get("CN MATCHED", 0),
        "Amendment Match":               status_counts.get("AMENDMENT MATCH", 0),
        "Probable Match":                status_counts.get("PROBABLE MATCH", 0),
        "Mismatch":                      status_counts.get("MISMATCH", 0),
        "Only in GSTR-2B (count)":       status_counts.get("ONLY IN GSTR-2B", 0),
        "Only in Books (count)":         status_counts.get("ONLY IN BOOKS", 0),
        "Only in GSTR-2B (₹)":          round(result_df[result_df["Status"]=="ONLY IN GSTR-2B"]["GSTR2B_Total_Tax"].sum(), 2),
        "Only in Books (₹)":            round(result_df[result_df["Status"]=="ONLY IN BOOKS"]["Books_Total_Tax"].sum(), 2),
    }

    return {"overall": overall, "category": category_df}


def build_vendor_summary(result_df: pd.DataFrame) -> pd.DataFrame:
    """Vendor-wise ITC comparison table."""
    grp = result_df.groupby("GSTIN", as_index=False).agg(
        Supplier_Name=("Supplier_Name", lambda x: x.dropna().iloc[0] if len(x.dropna()) else ""),
        Books_ITC=("Books_Total_Tax", "sum"),
        GSTR2B_ITC=("GSTR2B_Total_Tax", "sum"),
    )
    grp["Difference"] = grp["Books_ITC"] - grp["GSTR2B_ITC"]
    grp = grp.sort_values("Difference", ascending=False).reset_index(drop=True)
    return grp


def _safe_sum(df, col):
    return round(df[col].dropna().sum(), 2) if col in df.columns else 0.0


# ─────────────────────────────────────────────────────────────────────
# 6.  FILE READER (handles xlsx, xls, csv)
# ─────────────────────────────────────────────────────────────────────

def read_file(filepath: str) -> pd.DataFrame:
    """
    Read xlsx / xls / csv into DataFrame.
    Auto-detects header row (searches first 10 rows for GSTIN keyword).
    """
    ext = str(filepath).lower().rsplit(".", 1)[-1]

    if ext == "csv":
        df = pd.read_csv(filepath, dtype=str)
    else:
        # Try finding header row
        raw = pd.read_excel(filepath, header=None, nrows=15, dtype=str)
        header_row = 0
        for i, row in raw.iterrows():
            row_lower = " ".join(str(v).lower() for v in row.values)
            if any(k in row_lower for k in ["gstin", "invoice", "taxable", "igst"]):
                header_row = i
                break
        df = pd.read_excel(filepath, header=header_row, dtype=str)

    # Drop completely empty rows/cols
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df
