"""
matcher.py - Indian GST Multi-Factor Reconciliation Engine
GSTR2B Match Application
=========================================================
Implements deterministic and fuzzy multi-factor reconciliation between
Internal Purchase Books and GSTR-2B data (JSON, Excel, CSV, PDF).

Categorizes every record into:
1. Exact Match
2. Fuzzy Match (with discrepancy details)
3. Mismatch (GSTIN/Invoice match, but amount/date exceeds tolerance)
4. Missing in GSTR-2B (Blocked ITC risk)
5. Missing in Purchase Book (Unrecorded purchase / excess ITC on portal)
"""

import re
import json
from datetime import datetime
from typing import Tuple, List, Dict, Any, Optional
import pandas as pd
import numpy as np
from thefuzz import fuzz


COLUMN_ALIASES = {
    "GSTIN": ["gstin", "supplier gstin", "gstin/uin", "gstin of supplier", "vendor gstin", "party gstin"],
    "Vendor_Name": ["vendor name", "supplier name", "party name", "trade name", "trade/legal name", "legal name"],
    "Invoice_Number": [
        "invoice number", "invoice no", "invoice no.", "inv no", "inv no.",
        "document number", "doc no", "doc no.", "note number", "bill no"
    ],
    "Invoice_Date": ["invoice date", "inv date", "document date", "doc date", "date", "bill date"],
    "Taxable_Value": ["taxable value", "taxable value (₹)", "taxable amount", "taxable amt", "taxable val"],
    "IGST": ["igst", "igst amount", "integrated tax", "integrated tax (₹)", "igst (₹)"],
    "CGST": ["cgst", "cgst amount", "central tax", "central tax (₹)", "cgst (₹)"],
    "SGST": ["sgst", "sgst amount", "state tax", "state/ut tax", "state/ut tax (₹)", "sgst (₹)"],
    "CESS": ["cess", "cess amount", "cess (₹)"],
    "Total_Tax": ["total tax", "tax amount", "total tax amount", "total gst", "tax val"],
    "Document_Type": ["document type", "doc type", "invoice type", "note type", "type"],
    "Original_Invoice_Number": [
        "original invoice number", "original invoice no", "original document number", "orig inv no"
    ],
    "Original_Invoice_Date": [
        "original invoice date", "original document date", "orig inv date"
    ],
    "Reverse_Charge": [
        "reverse charge", "rcm", "is reverse charge", "reverse charge applicable", "supply attract reverse charge"
    ],
}

FINANCIAL_COLS = ["Taxable_Value", "IGST", "CGST", "SGST", "CESS", "Total_Tax"]


def clean_invoice_number(val: Any) -> str:
    """
    Standardizes invoice numbers:
    - Strips special characters (/ - _ \ spaces)
    - Strips leading zeros (e.g. '00123' -> '123')
    - Converts to uppercase
    """
    if pd.isna(val) or str(val).strip() == "":
        return "UNKNOWN_INV"
    # Remove all non-alphanumeric characters
    cleaned = re.sub(r"[^A-Z0-9]", "", str(val).upper().strip())
    # Strip leading zeros for numeric sequences while retaining the alphanumeric stem
    cleaned = re.sub(r"^0+", "", cleaned)
    return cleaned if cleaned else "0"


def normalize_gstin(val: Any) -> str:
    """Normalizes GSTIN string: uppercase, no spaces."""
    if pd.isna(val) or str(val).strip() == "":
        return "UNKNOWN_GSTIN"
    return re.sub(r"\s+", "", str(val).upper().strip())


def validate_gstin_format(val: str) -> bool:
    """Validates 15-character Indian GSTIN pattern."""
    if not val or not isinstance(val, str):
        return False
    pattern = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
    return bool(re.fullmatch(pattern, val.strip().upper()))


def parse_date_safe(val: Any) -> pd.Timestamp:
    """Safely converts string or date object to pandas Timestamp."""
    if pd.isna(val) or str(val).strip() == "":
        return pd.NaT
    return pd.to_datetime(val, dayfirst=True, errors="coerce")


def infer_doc_type(source_section: str = "", raw_type: str = "") -> str:
    """Infers standard document type (INVOICE, CREDIT_NOTE, DEBIT_NOTE, AMENDMENT)."""
    text = f"{source_section} {raw_type}".upper().strip()
    
    if any(k in text for k in ["CREDIT", " CR", "CDN", "C"]):
        if "AMEND" in text:
            return "AMENDED_CREDIT_NOTE"
        return "CREDIT_NOTE"
    if any(k in text for k in ["DEBIT", " DR", "DN", "D"]):
        if "AMEND" in text:
            return "AMENDED_DEBIT_NOTE"
        return "DEBIT_NOTE"
    if "AMEND" in text or "B2BA" in text:
        return "AMENDMENT"
    return "INVOICE"


def map_dataframe_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str], List[str]]:
    """
    Standardizes DataFrame column names by matching against COLUMN_ALIASES.
    Returns: (mapped_df, ambiguous_columns, missing_columns)
    """
    mapped_df = df.copy()
    col_mapping = {}
    ambiguous = []
    
    # Lowercase & strip existing column names
    col_lookup = {str(col).strip().lower(): col for col in mapped_df.columns}

    for std_name, aliases in COLUMN_ALIASES.items():
        found = [col_lookup[a] for a in aliases if a in col_lookup]
        if len(found) == 1:
            col_mapping[found[0]] = std_name
        elif len(found) > 1:
            ambiguous.append(std_name)
            col_mapping[found[0]] = std_name  # Pick the first match

    mapped_df = mapped_df.rename(columns=col_mapping)

    missing = []
    for std_name in COLUMN_ALIASES:
        if std_name not in mapped_df.columns:
            mapped_df[std_name] = np.nan
            missing.append(std_name)

    return mapped_df, ambiguous, missing


def preprocess_data(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """
    Preprocesses Purchase Book or GSTR-2B DataFrame:
    - Standardizes columns
    - Parses GSTIN, Invoice Number, Dates, and Financials
    - Flags RCM invoices
    - Flags duplicates and data quality issues
    """
    df_clean, ambiguous, missing_cols = map_dataframe_columns(df)

    # Retain raw values for reports and audit
    for col in COLUMN_ALIASES:
        df_clean[f"Raw_{col}"] = df_clean[col]

    df_clean["Parsed_GSTIN"] = df_clean["GSTIN"].apply(normalize_gstin)
    df_clean["Parsed_Inv"] = df_clean["Invoice_Number"].apply(clean_invoice_number)
    df_clean["Parsed_Date"] = df_clean["Invoice_Date"].apply(parse_date_safe)
    
    # Document Type & Original Invoice for Credit/Debit Notes
    df_clean["Parsed_Document_Type"] = [
        infer_doc_type(
            str(df_clean.iloc[i].get("Source_Section", "")),
            str(df_clean.iloc[i].get("Document_Type", ""))
        )
        for i in range(len(df_clean))
    ]
    df_clean["Parsed_Original_Inv"] = df_clean["Original_Invoice_Number"].apply(clean_invoice_number)

    # RCM (Reverse Charge Mechanism) Isolation
    def parse_rcm_flag(v: Any) -> str:
        if pd.isna(v):
            return "NO"
        s = str(v).strip().upper()
        return "YES" if s in ["Y", "YES", "TRUE", "1"] else "NO"

    df_clean["Parsed_RCM"] = df_clean["Reverse_Charge"].apply(parse_rcm_flag)

    # Clean numeric fields
    for col in FINANCIAL_COLS:
        # Convert strings with commas or currency symbols
        if df_clean[col].dtype == object:
            df_clean[f"Parsed_{col}"] = (
                df_clean[col].astype(str)
                .str.replace(r"[₹Rs\s,]", "", regex=True)
            )
            df_clean[f"Parsed_{col}"] = pd.to_numeric(df_clean[f"Parsed_{col}"], errors="coerce").fillna(0.0)
        else:
            df_clean[f"Parsed_{col}"] = pd.to_numeric(df_clean[col], errors="coerce").fillna(0.0)

    # If Total_Tax is missing or zero, compute it
    comp_tax = df_clean["Parsed_IGST"] + df_clean["Parsed_CGST"] + df_clean["Parsed_SGST"] + df_clean["Parsed_CESS"]
    df_clean["Parsed_Total_Tax"] = np.where(
        df_clean["Parsed_Total_Tax"] > 0,
        df_clean["Parsed_Total_Tax"],
        comp_tax
    )

    df_clean["Data_Quality_Issue"] = ""
    df_clean["Match_Category"] = "PENDING"
    df_clean["Row_ID"] = [f"{source_name}_{i+1}" for i in range(len(df_clean))]

    # Flag Invalid GSTINs
    invalid_gstin = ~df_clean["Parsed_GSTIN"].apply(validate_gstin_format) & (df_clean["Parsed_GSTIN"] != "UNKNOWN_GSTIN")
    df_clean.loc[invalid_gstin, "Data_Quality_Issue"] += "Invalid GSTIN format; "

    # Flag missing keys
    df_clean.loc[df_clean["Parsed_GSTIN"] == "UNKNOWN_GSTIN", "Data_Quality_Issue"] += "Missing GSTIN; "
    df_clean.loc[df_clean["Parsed_Inv"] == "UNKNOWN_INV", "Data_Quality_Issue"] += "Missing Invoice No; "

    # Flag internal duplicates
    valid_keys = (df_clean["Parsed_GSTIN"] != "UNKNOWN_GSTIN") & (df_clean["Parsed_Inv"] != "UNKNOWN_INV")
    dupes = df_clean.duplicated(subset=["Parsed_GSTIN", "Parsed_Inv"], keep=False) & valid_keys
    df_clean.loc[dupes, "Data_Quality_Issue"] += f"Duplicate in {source_name}; "

    return df_clean


def is_amount_within_tolerance(
    val1: float, 
    val2: float, 
    abs_tol: float = 1.0, 
    rel_tol_pct: float = 1.0
) -> Tuple[bool, float]:
    """
    Checks if two amounts match within configurable tolerance:
    - Absolute difference (e.g. ±₹1.00) OR
    - Relative percentage difference (e.g. ±1.0%)
    """
    diff = abs(val1 - val2)
    if diff <= abs_tol:
        return True, diff
    base = max(abs(val1), abs(val2), 0.01)
    pct_diff = (diff / base) * 100.0
    if pct_diff <= rel_tol_pct:
        return True, diff
    return False, diff


def parse_gstr2b_json(json_bytes: bytes) -> Tuple[pd.DataFrame, List[str]]:
    """
    Robust JSON parser for official GST portal GSTR-2B downloads.
    Extracts b2b, b2ba, cdnr, and cdnra sections.
    """
    warnings = []
    try:
        raw_text = json_bytes.decode("utf-8", errors="replace")
        data = json.loads(raw_text)
    except Exception as e:
        raise ValueError(f"Failed to read JSON: {str(e)}")

    rows = []
    
    # Recursive search for B2B/CDNR tables in GSTR-2B JSON hierarchy
    def search_sections(node, section_name):
        results = []
        if isinstance(node, dict):
            for k, v in node.items():
                if str(k).lower() == section_name.lower() and isinstance(v, list):
                    results.append(v)
                else:
                    results.extend(search_sections(v, section_name))
        elif isinstance(node, list):
            for item in node:
                results.extend(search_sections(item, section_name))
        return results

    sections_to_check = {
        "b2b": "B2B",
        "b2ba": "B2B_AMENDED",
        "cdnr": "CDNR",
        "cdnra": "CDNR_AMENDED",
    }

    for sec_key, sec_label in sections_to_check.items():
        found_lists = search_sections(data, sec_key)
        for sec_list in found_lists:
            for supplier in sec_list:
                ctin = supplier.get("ctin", "")
                trade_name = supplier.get("trdnm", supplier.get("trdname", ""))
                
                doc_list = supplier.get("inv") or supplier.get("nt") or []
                for doc in doc_list:
                    inum = doc.get("inum", doc.get("ntnum", ""))
                    idt = doc.get("idt", doc.get("ntdt", ""))
                    doc_type = doc.get("typ", doc.get("ntty", doc.get("type", "")))
                    orig_inv = doc.get("oinum", doc.get("ontnum", ""))
                    rcm_flag = doc.get("rchrg", doc.get("rev", "N"))
                    
                    # Items list for tax breakdown
                    items = doc.get("itms") or doc.get("items") or []
                    
                    def get_float(d, *keys):
                        for k in keys:
                            val = d.get(k)
                            if val is not None and str(val).strip() != "":
                                try:
                                    return float(val)
                                except ValueError:
                                    pass
                        return 0.0

                    txval = 0.0
                    igst = 0.0
                    cgst = 0.0
                    sgst = 0.0
                    cess = 0.0

                    if items:
                        for it in items:
                            item_det = it.get("itm_det", it)
                            txval += get_float(item_det, "txval")
                            igst += get_float(item_det, "iamt", "igst")
                            cgst += get_float(item_det, "camt", "cgst")
                            sgst += get_float(item_det, "samt", "sgst")
                            cess += get_float(item_det, "csamt", "cess")
                    else:
                        txval = get_float(doc, "txval", "val")
                        igst = get_float(doc, "iamt", "igst")
                        cgst = get_float(doc, "camt", "cgst")
                        sgst = get_float(doc, "samt", "sgst")
                        cess = get_float(doc, "csamt", "cess")

                    rows.append({
                        "GSTIN": ctin,
                        "Vendor_Name": trade_name,
                        "Invoice_Number": inum,
                        "Invoice_Date": idt,
                        "Taxable_Value": txval,
                        "IGST": igst,
                        "CGST": cgst,
                        "SGST": sgst,
                        "CESS": cess,
                        "Total_Tax": igst + cgst + sgst + cess,
                        "Document_Type": doc_type,
                        "Original_Invoice_Number": orig_inv,
                        "Reverse_Charge": rcm_flag,
                        "Source_Section": sec_label
                    })

    if not rows:
        warnings.append("No standard B2B or CDNR invoice sections found in this GSTR-2B JSON.")
        return pd.DataFrame(), warnings

    return pd.DataFrame(rows), warnings


def perform_reconciliation(
    books_df: pd.DataFrame,
    gstr2b_df: pd.DataFrame,
    abs_tolerance: float = 1.0,
    rel_tolerance_pct: float = 1.0,
    date_tolerance_days: int = 3,
    min_fuzzy_inv_score: int = 75
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """
    Executes the multi-factor deterministic and fuzzy reconciliation engine.
    Returns:
    - categorized_books: Books DataFrame with category and matching details
    - missing_in_books: GSTR-2B rows not present in Purchase Book
    - all_gstr2b: Processed GSTR-2B records
    - summary_metrics: Key ITC reconciliation statistics
    """
    books = preprocess_data(books_df, "BOOKS")
    gstr2b = preprocess_data(gstr2b_df, "GSTR2B")

    # Tracking fields
    books["Match_Category"] = "Missing in GSTR-2B"
    books["Matched_2B_Invoice"] = ""
    books["Matched_2B_RowID"] = ""
    books["Discrepancy_Note"] = ""
    books["Match_Confidence"] = 0.0

    matched_2b_row_ids = set()

    # Exclude RCM from standard ITC matching pool (RCM is paid via cash ledger, not standard ITC)
    non_rcm_books = books[books["Parsed_RCM"] != "YES"].copy()
    non_rcm_2b = gstr2b[gstr2b["Parsed_RCM"] != "YES"].copy()

    # -------------------------------------------------------------------------
    # PASS 1: DETERMINISTIC EXACT MATCH (GSTIN + Normalized Invoice Number)
    # -------------------------------------------------------------------------
    for b_idx, b_row in non_rcm_books.iterrows():
        b_gstin = b_row["Parsed_GSTIN"]
        b_inv = b_row["Parsed_Inv"]

        if b_gstin == "UNKNOWN_GSTIN" or b_inv == "UNKNOWN_INV":
            continue

        # Look for identical GSTIN and cleaned invoice
        candidates = non_rcm_2b[
            (non_rcm_2b["Parsed_GSTIN"] == b_gstin) &
            (non_rcm_2b["Parsed_Inv"] == b_inv) &
            (~non_rcm_2b["Row_ID"].isin(matched_2b_row_ids))
        ]

        if candidates.empty:
            continue

        match_row = candidates.iloc[0]
        
        # Check financial tolerances
        tax_matched, tax_diff = is_amount_within_tolerance(
            b_row["Parsed_Total_Tax"], match_row["Parsed_Total_Tax"], abs_tolerance, rel_tolerance_pct
        )
        taxable_matched, taxable_diff = is_amount_within_tolerance(
            b_row["Parsed_Taxable_Value"], match_row["Parsed_Taxable_Value"], abs_tolerance, rel_tolerance_pct
        )

        # Check date tolerance
        date_matched = True
        date_diff_days = 0
        if pd.notna(b_row["Parsed_Date"]) and pd.notna(match_row["Parsed_Date"]):
            date_diff_days = abs((b_row["Parsed_Date"] - match_row["Parsed_Date"]).days)
            date_matched = date_diff_days <= date_tolerance_days

        if tax_matched and taxable_matched and date_matched:
            # Exact Match
            books.at[b_idx, "Match_Category"] = "Exact Match"
            books.at[b_idx, "Matched_2B_Invoice"] = match_row["Raw_Invoice_Number"]
            books.at[b_idx, "Matched_2B_RowID"] = match_row["Row_ID"]
            books.at[b_idx, "Match_Confidence"] = 100.0
            books.at[b_idx, "Discrepancy_Note"] = "Perfect match on GSTIN, Invoice, Tax, and Date"
            matched_2b_row_ids.add(match_row["Row_ID"])
        else:
            # Mismatch (both exist but amount or date differs beyond tolerance)
            discrepancy_reasons = []
            if not tax_matched:
                discrepancy_reasons.append(f"Tax diff ₹{tax_diff:.2f}")
            if not taxable_matched:
                discrepancy_reasons.append(f"Taxable diff ₹{taxable_diff:.2f}")
            if not date_matched:
                discrepancy_reasons.append(f"Date diff {date_diff_days} days")

            books.at[b_idx, "Match_Category"] = "Mismatch"
            books.at[b_idx, "Matched_2B_Invoice"] = match_row["Raw_Invoice_Number"]
            books.at[b_idx, "Matched_2B_RowID"] = match_row["Row_ID"]
            books.at[b_idx, "Match_Confidence"] = 85.0
            books.at[b_idx, "Discrepancy_Note"] = " | ".join(discrepancy_reasons)
            matched_2b_row_ids.add(match_row["Row_ID"])

    # -------------------------------------------------------------------------
    # PASS 2: FUZZY MATCHING (Typo in invoice number or formatting difference)
    # -------------------------------------------------------------------------
    unmatched_books = books[
        (books["Match_Category"] == "Missing in GSTR-2B") &
        (books["Parsed_RCM"] != "YES") &
        (books["Parsed_GSTIN"] != "UNKNOWN_GSTIN")
    ]

    for b_idx, b_row in unmatched_books.iterrows():
        b_gstin = b_row["Parsed_GSTIN"]
        b_inv = b_row["Parsed_Inv"]

        # Search candidates from same GSTIN that are still unmatched
        same_vendor_2b = non_rcm_2b[
            (non_rcm_2b["Parsed_GSTIN"] == b_gstin) &
            (~non_rcm_2b["Row_ID"].isin(matched_2b_row_ids))
        ]

        if same_vendor_2b.empty:
            continue

        best_candidate = None
        highest_inv_sim = 0

        for _, c_row in same_vendor_2b.iterrows():
            c_inv = c_row["Parsed_Inv"]
            sim_score = fuzz.ratio(b_inv, c_inv)

            if sim_score >= min_fuzzy_inv_score and sim_score > highest_inv_sim:
                # Test financial compatibility
                tax_matched, tax_diff = is_amount_within_tolerance(
                    b_row["Parsed_Total_Tax"], c_row["Parsed_Total_Tax"], abs_tolerance, rel_tolerance_pct
                )
                if tax_matched:
                    best_candidate = (c_row, sim_score, tax_diff)
                    highest_inv_sim = sim_score

        if best_candidate:
            c_row, inv_sim, tax_diff = best_candidate
            books.at[b_idx, "Match_Category"] = "Fuzzy Match"
            books.at[b_idx, "Matched_2B_Invoice"] = c_row["Raw_Invoice_Number"]
            books.at[b_idx, "Matched_2B_RowID"] = c_row["Row_ID"]
            books.at[b_idx, "Match_Confidence"] = float(inv_sim)
            books.at[b_idx, "Discrepancy_Note"] = (
                f"Fuzzy Invoice Similarity: {inv_sim}% (Books: '{b_row['Raw_Invoice_Number']}' vs 2B: '{c_row['Raw_Invoice_Number']}'); "
                f"Tax diff: ₹{tax_diff:.2f}"
            )
            matched_2b_row_ids.add(c_row["Row_ID"])

    # Flag RCM Invoices distinctly
    books.loc[books["Parsed_RCM"] == "YES", "Match_Category"] = "RCM Invoice (Isolated)"
    books.loc[books["Parsed_RCM"] == "YES", "Discrepancy_Note"] = "Reverse Charge Mechanism applicable. Pay via Cash ledger; cannot reconcile against regular ITC."

    # -------------------------------------------------------------------------
    # PASS 3: UNMATCHED GSTR-2B (Missing in Purchase Book)
    # -------------------------------------------------------------------------
    missing_in_books = gstr2b[
        (~gstr2b["Row_ID"].isin(matched_2b_row_ids)) &
        (gstr2b["Parsed_RCM"] != "YES")
    ].copy()
    missing_in_books["Match_Category"] = "Missing in Purchase Book"
    missing_in_books["Discrepancy_Note"] = "Available on GST portal but unrecorded in internal Purchase Book. Verify invoice receipt."

    # -------------------------------------------------------------------------
    # SUMMARY METRICS CALCULATION
    # -------------------------------------------------------------------------
    cat_counts = books["Match_Category"].value_counts().to_dict()
    
    total_itc_claimed = books.loc[
        books["Match_Category"].isin(["Exact Match", "Fuzzy Match"]), "Parsed_Total_Tax"
    ].sum()

    total_itc_blocked = books.loc[
        books["Match_Category"] == "Missing in GSTR-2B", "Parsed_Total_Tax"
    ].sum()

    total_mismatch_itc = books.loc[
        books["Match_Category"] == "Mismatch", "Parsed_Total_Tax"
    ].sum()

    unrecorded_2b_itc = missing_in_books["Parsed_Total_Tax"].sum()

    summary_metrics = {
        "total_books_rows": len(books),
        "total_gstr2b_rows": len(gstr2b),
        "exact_matches": cat_counts.get("Exact Match", 0),
        "fuzzy_matches": cat_counts.get("Fuzzy Match", 0),
        "mismatches": cat_counts.get("Mismatch", 0),
        "missing_in_2b": cat_counts.get("Missing in GSTR-2B", 0),
        "missing_in_books": len(missing_in_books),
        "rcm_count": cat_counts.get("RCM Invoice (Isolated)", 0),
        "total_itc_claimed": total_itc_claimed,
        "total_itc_blocked_risk": total_itc_blocked,
        "total_mismatch_itc": total_mismatch_itc,
        "unrecorded_2b_itc": unrecorded_2b_itc,
    }

    return books, missing_in_books, gstr2b, summary_metrics
