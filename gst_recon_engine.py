"""
GST Reconciliation Engine v2
==============================
Supports both GSTR-2A and GSTR-2B formats auto-detected on upload.

GSTR-2B Format: GSTIN, Invoice No, Invoice Date, Taxable Value, IGST, CGST, SGST
GSTR-2A Format: GSTIN, SUPPLIER NAME, INVOICE NUBER, TAXABLE VALUE, IGST-2A, CGST-2A, SGST-2A
Books Format  : GSTIN/UIN, Particulars, Voucher No., Date, IGST, CGST, SGST

MATCHING LOGIC:
  Since Books use internal Voucher Numbers (not supplier invoice numbers),
  invoice-level matching is not possible. Matching is done at:
  Level 1 — GSTIN + Total Tax (diff <= tolerance) + Date within +/-5 days  -> MATCHED
  Level 2 — GSTIN + Total Tax (diff <= tolerance) no date restriction       -> PROBABLE MATCH
  Level 3 — GSTIN in both but amounts differ                                -> VENDOR MISMATCH
  Level 4 — Only in 2A/2B or Only in Books
"""

import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

TOLERANCE   = 10.0
DATE_WINDOW = 5


def detect_file_type(path):
    for header_row in [0, 1]:
        try:
            df   = pd.read_excel(path, nrows=5, header=header_row)
            cols = " ".join([str(c).strip().upper() for c in df.columns])
            if "IGST - 2A" in cols or "CGST - 2A" in cols or "INVOICE NUBER" in cols:
                return "2A"
            if "GSTIN/UIN" in cols or "VOUCHER NO" in cols or "PARTICULARS" in cols:
                return "BOOKS"
            if "SUPPLIER GSTIN" in cols or "TRADE NAME" in cols or "LEGAL NAME" in cols:
                return "2B"
            if "IGST" in cols and "GSTIN" in cols and "TAXABLE" in cols:
                return "2A"
            if "IGST" in cols and "GSTIN" in cols:
                return "BOOKS"
        except Exception:
            continue
    return "UNKNOWN"


def read_2a(path):
    df = pd.read_excel(path, header=1)
    df = df.rename(columns={
        "GSTIN":         "GSTIN",
        "SUPPLIER NAME": "Supplier_Name",
        "INVOICE NUBER": "Invoice_No",
        "TAXABLE VALUE": "Taxable_Value",
        "IGST - 2A":     "IGST",
        "CGST - 2A":     "CGST",
        "SGST - 2A":     "SGST",
        "Date":          "Date",
    })
    for c in ["IGST","CGST","SGST","Taxable_Value"]:
        if c not in df.columns: df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Total_Tax"]  = df["IGST"] + df["CGST"] + df["SGST"]
    df["GSTIN"]      = df["GSTIN"].astype(str).str.strip().str.upper()
    df["Date"]       = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)
    df["Invoice_No"] = df.get("Invoice_No", pd.Series([""] * len(df))).astype(str).str.strip()
    if "Supplier_Name" not in df.columns: df["Supplier_Name"] = ""
    if "Taxable_Value" not in df.columns: df["Taxable_Value"] = 0
    df["Source"] = "GSTR-2A"
    df = df[df["GSTIN"].str.len() == 15].reset_index(drop=True)
    return df


def _find_col(cols_upper, options):
    for opt in options:
        if opt in cols_upper:
            return cols_upper[opt]
    return None


def read_2b(path):
    for header_row in [0, 1, 2]:
        try:
            df = pd.read_excel(path, header=header_row)
            cols_upper = {str(c).strip().upper(): c for c in df.columns}
            gstin_col = _find_col(cols_upper, ["GSTIN","SUPPLIER GSTIN","GSTIN OF SUPPLIER","VENDOR GSTIN"])
            inv_col   = _find_col(cols_upper, ["INVOICE NUMBER","INVOICE NO","INVOICE NO.","INV NO","DOCUMENT NO"])
            date_col  = _find_col(cols_upper, ["INVOICE DATE","DATE","INVOICE DT","DOC DATE"])
            tax_col   = _find_col(cols_upper, ["TAXABLE VALUE","TAXABLE AMOUNT","TAXABLE"])
            igst_col  = _find_col(cols_upper, ["IGST","IGST AMOUNT"])
            cgst_col  = _find_col(cols_upper, ["CGST","CGST AMOUNT"])
            sgst_col  = _find_col(cols_upper, ["SGST","SGST AMOUNT","SGST/UTGST"])
            name_col  = _find_col(cols_upper, ["TRADE NAME","SUPPLIER NAME","LEGAL NAME","NAME"])
            if not gstin_col:
                continue
            rename_map = {}
            if gstin_col: rename_map[gstin_col] = "GSTIN"
            if inv_col:   rename_map[inv_col]   = "Invoice_No"
            if date_col:  rename_map[date_col]  = "Date"
            if tax_col:   rename_map[tax_col]   = "Taxable_Value"
            if igst_col:  rename_map[igst_col]  = "IGST"
            if cgst_col:  rename_map[cgst_col]  = "CGST"
            if sgst_col:  rename_map[sgst_col]  = "SGST"
            if name_col:  rename_map[name_col]  = "Supplier_Name"
            df = df.rename(columns=rename_map)
            for c in ["IGST","CGST","SGST","Taxable_Value"]:
                if c not in df.columns: df[c] = 0
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
            df["Total_Tax"]  = df["IGST"] + df["CGST"] + df["SGST"]
            df["GSTIN"]      = df["GSTIN"].astype(str).str.strip().str.upper()
            df["Invoice_No"] = df.get("Invoice_No", pd.Series([""] * len(df))).astype(str).str.strip()
            df["Date"]       = pd.to_datetime(df.get("Date", None), errors="coerce", dayfirst=True) if "Date" in df.columns else pd.NaT
            if "Supplier_Name" not in df.columns: df["Supplier_Name"] = ""
            df["Source"] = "GSTR-2B"
            df = df[df["GSTIN"].str.len() == 15].reset_index(drop=True)
            return df
        except Exception:
            continue
    raise ValueError("Could not read GSTR-2B file. Please check the format.")


def read_books(path):
    df = pd.read_excel(path, header=1)
    df = df.rename(columns={
        "GSTIN/UIN":   "GSTIN",
        "Particulars": "Supplier_Name",
        "Voucher No.": "Voucher_No",
        "Date":        "Date",
        "IGST":        "IGST",
        "CGST":        "CGST",
        "SGST":        "SGST",
    })
    if "GSTIN" not in df.columns:
        for col in df.columns:
            if "GSTIN" in str(col).upper():
                df = df.rename(columns={col: "GSTIN"})
                break
    for c in ["IGST","CGST","SGST"]:
        if c not in df.columns: df[c] = 0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Total_Tax"]  = df["IGST"] + df["CGST"] + df["SGST"]
    df["GSTIN"]      = df["GSTIN"].astype(str).str.strip().str.upper()
    df["Date"]       = pd.to_datetime(df["Date"], errors="coerce")
    df["Voucher_No"] = df["Voucher_No"].astype(str).str.strip() if "Voucher_No" in df.columns else ""
    if "Supplier_Name" not in df.columns: df["Supplier_Name"] = ""
    df["Source"] = "BOOKS"
    df = df[df["GSTIN"].str.len() == 15].reset_index(drop=True)
    return df


def read_gst_file(path, file_type=None):
    if file_type is None:
        file_type = detect_file_type(path)
    if file_type == "2A":
        return read_2a(path), "GSTR-2A"
    elif file_type == "2B":
        return read_2b(path), "GSTR-2B"
    elif file_type == "BOOKS":
        return read_books(path), "Books"
    else:
        try:
            return read_2a(path), "GSTR-2A"
        except Exception:
            try:
                return read_2b(path), "GSTR-2B"
            except Exception:
                raise ValueError("Could not detect file format. Please upload GSTR-2A, GSTR-2B or Books file.")


def reconcile(df_gst, df_books, tolerance=TOLERANCE, date_window=DATE_WINDOW):
    results       = []
    matched_gst   = set()
    matched_books = set()
    source_label  = df_gst["Source"].iloc[0] if "Source" in df_gst.columns else "GSTR"

    # Level 1 — GSTIN + Tax + Date
    for i, rg in df_gst.iterrows():
        if i in matched_gst: continue
        cands = df_books[
            (df_books["GSTIN"] == rg["GSTIN"]) &
            (~df_books.index.isin(matched_books)) &
            (abs(df_books["Total_Tax"] - rg["Total_Tax"]) <= tolerance)
        ]
        if not cands.empty and pd.notna(rg.get("Date")):
            cands = cands[cands["Date"].apply(
                lambda d: abs((d - rg["Date"]).days) <= date_window if pd.notna(d) else True
            )]
        if not cands.empty:
            j = cands.index[0]
            results.append(_row(rg, df_books.loc[j], "MATCHED", source_label))
            matched_gst.add(i); matched_books.add(j)

    # Level 2 — GSTIN + Tax (no date)
    for i, rg in df_gst.iterrows():
        if i in matched_gst: continue
        cands = df_books[
            (df_books["GSTIN"] == rg["GSTIN"]) &
            (~df_books.index.isin(matched_books)) &
            (abs(df_books["Total_Tax"] - rg["Total_Tax"]) <= tolerance)
        ]
        if not cands.empty:
            j = cands.index[0]
            results.append(_row(rg, df_books.loc[j], "PROBABLE MATCH", source_label))
            matched_gst.add(i); matched_books.add(j)

    # Level 3 — GSTIN only (amount differs)
    for i, rg in df_gst.iterrows():
        if i in matched_gst: continue
        cands = df_books[
            (df_books["GSTIN"] == rg["GSTIN"]) &
            (~df_books.index.isin(matched_books))
        ]
        if not cands.empty:
            j = cands.index[0]
            results.append(_row(rg, df_books.loc[j], "VENDOR MISMATCH", source_label))
            matched_gst.add(i); matched_books.add(j)

    # Only in GST
    for i, rg in df_gst.iterrows():
        if i in matched_gst: continue
        results.append({
            "GSTIN": rg["GSTIN"], "Supplier_Name": rg.get("Supplier_Name",""),
            "Invoice_No": rg.get("Invoice_No",""),
            "Date_GST": rg.get("Date"), "Date_Books": None, "Voucher_No": None,
            "GST_IGST": rg.get("IGST",0), "GST_CGST": rg.get("CGST",0), "GST_SGST": rg.get("SGST",0),
            "GST_Taxable": rg.get("Taxable_Value",0), "GST_Total_Tax": rg.get("Total_Tax",0),
            "Books_IGST": None, "Books_CGST": None, "Books_SGST": None, "Books_Total_Tax": None,
            "Difference": None, "Status": f"ONLY IN {source_label.upper()}", "Source": source_label,
        })

    # Only in Books
    for j, rb in df_books.iterrows():
        if j in matched_books: continue
        results.append({
            "GSTIN": rb["GSTIN"], "Supplier_Name": rb.get("Supplier_Name",""),
            "Invoice_No": None,
            "Date_GST": None, "Date_Books": rb.get("Date"), "Voucher_No": rb.get("Voucher_No",""),
            "GST_IGST": None, "GST_CGST": None, "GST_SGST": None,
            "GST_Taxable": None, "GST_Total_Tax": None,
            "Books_IGST": rb.get("IGST",0), "Books_CGST": rb.get("CGST",0), "Books_SGST": rb.get("SGST",0),
            "Books_Total_Tax": rb.get("Total_Tax",0),
            "Difference": None, "Status": "ONLY IN BOOKS", "Source": source_label,
        })

    return pd.DataFrame(results)


def _row(rg, rb, status, source_label):
    diff = round(rb.get("Total_Tax",0) - rg.get("Total_Tax",0), 2)
    return {
        "GSTIN": rg["GSTIN"],
        "Supplier_Name": rg.get("Supplier_Name") or rb.get("Supplier_Name",""),
        "Invoice_No": rg.get("Invoice_No",""),
        "Date_GST": rg.get("Date"), "Date_Books": rb.get("Date"),
        "Voucher_No": rb.get("Voucher_No",""),
        "GST_IGST": rg.get("IGST",0), "GST_CGST": rg.get("CGST",0), "GST_SGST": rg.get("SGST",0),
        "GST_Taxable": rg.get("Taxable_Value",0), "GST_Total_Tax": rg.get("Total_Tax",0),
        "Books_IGST": rb.get("IGST",0), "Books_CGST": rb.get("CGST",0), "Books_SGST": rb.get("SGST",0),
        "Books_Total_Tax": rb.get("Total_Tax",0),
        "Difference": diff, "Status": status, "Source": source_label,
    }


def build_summary(result_df, df_gst, df_books):
    sc           = result_df["Status"].value_counts().to_dict()
    total_gst    = df_gst["Total_Tax"].sum()
    total_books  = df_books["Total_Tax"].sum()
    source_label = df_gst["Source"].iloc[0] if "Source" in df_gst.columns else "GSTR"
    only_gst_key = f"ONLY IN {source_label.upper()}"
    only_gst_amt = result_df[result_df["Status"]==only_gst_key]["GST_Total_Tax"].sum()
    only_bk_amt  = result_df[result_df["Status"]=="ONLY IN BOOKS"]["Books_Total_Tax"].sum()
    return {
        f"Total ITC as per {source_label} (Rs)": round(total_gst, 2),
        "Total ITC as per Books (Rs)":            round(total_books, 2),
        "Net Difference (Rs)":                    round(total_books - total_gst, 2),
        "Matched (count)":                        sc.get("MATCHED", 0),
        "Probable Match (count)":                 sc.get("PROBABLE MATCH", 0),
        "Vendor Mismatch (count)":                sc.get("VENDOR MISMATCH", 0),
        f"Only in {source_label} (count)":        sc.get(only_gst_key, 0),
        f"Only in {source_label} - ITC (Rs)":     round(only_gst_amt, 2),
        "Only in Books (count)":                  sc.get("ONLY IN BOOKS", 0),
        "Only in Books - ITC (Rs)":               round(only_bk_amt, 2),
    }


def build_vendor_summary(result_df, source_label="GSTR"):
    rows = []
    for gstin, grp in result_df.groupby("GSTIN"):
        name  = grp["Supplier_Name"].dropna()
        name  = name.iloc[0] if len(name) else ""
        t_gst = grp["GST_Total_Tax"].sum()
        t_bks = grp["Books_Total_Tax"].sum()
        diff  = round(t_bks - t_gst, 2)
        status = grp["Status"].mode()[0] if len(grp) else ""
        rows.append({"GSTIN": gstin, "Supplier_Name": name,
                     "ITC_GST": round(t_gst,2), "ITC_Books": round(t_bks,2),
                     "Difference": diff, "Status": status})
    return pd.DataFrame(rows).sort_values("Difference").reset_index(drop=True)
