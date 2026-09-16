"""
db.py - SQLite Ledger for Cross-Month Carry-Forward ITC Tracking
GSTR2B Match Application
================================================================
This module maintains a local SQLite database so that invoices that were
previously marked as "Missing in GSTR-2B" in an earlier month can be
automatically matched when the vendor files late in a subsequent month.
"""

import sqlite3
import pandas as pd
from datetime import datetime
from typing import List, Dict, Any, Optional

DEFAULT_DB_PATH = "gstr2b_carry_forward.db"


def get_connection(db_path: str = DEFAULT_DB_PATH) -> sqlite3.Connection:
    """Returns a SQLite connection with row factory enabled."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_database(db_path: str = DEFAULT_DB_PATH) -> None:
    """
    Initializes the SQLite schema if it does not already exist.
    Stores un-reconciled purchase invoices carried forward from past months.
    """
    conn = get_connection(db_path)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pending_itc_ledger (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_label TEXT NOT NULL,
            gstin TEXT NOT NULL,
            clean_invoice_number TEXT NOT NULL,
            raw_invoice_number TEXT NOT NULL,
            invoice_date TEXT,
            vendor_name TEXT,
            taxable_value REAL,
            igst REAL,
            cgst REAL,
            sgst REAL,
            cess REAL,
            total_tax REAL,
            doc_type TEXT DEFAULT 'INVOICE',
            status TEXT DEFAULT 'PENDING',
            saved_at TEXT NOT NULL,
            resolved_period TEXT,
            resolved_at TEXT,
            UNIQUE(period_label, gstin, clean_invoice_number)
        );
    """)
    cursor.execute("""
        CREATE INDEX IF NOT EXISTS idx_pending_gstin_inv 
        ON pending_itc_ledger(gstin, clean_invoice_number, status);
    """)
    conn.commit()
    conn.close()


def save_missing_invoices(
    missing_df: pd.DataFrame, 
    period_label: str, 
    db_path: str = DEFAULT_DB_PATH
) -> int:
    """
    Saves invoices marked as 'Missing in GSTR-2B' into the database.
    Ignores rows that were already saved for the same period to prevent duplicates.
    Returns the count of newly saved records.
    """
    if missing_df.empty:
        return 0

    init_database(db_path)
    conn = get_connection(db_path)
    cursor = conn.cursor()
    saved_count = 0
    now_str = datetime.now().isoformat()

    for _, row in missing_df.iterrows():
        clean_inv = str(row.get("Parsed_Inv", "")).strip()
        gstin = str(row.get("Parsed_GSTIN", "")).strip()

        # Do not save if invoice or GSTIN is missing/invalid
        if not clean_inv or clean_inv == "UNKNOWN_INV" or not gstin or gstin == "UNKNOWN_GSTIN":
            continue

        try:
            cursor.execute("""
                INSERT OR IGNORE INTO pending_itc_ledger (
                    period_label, gstin, clean_invoice_number, raw_invoice_number,
                    invoice_date, vendor_name, taxable_value, igst, cgst, sgst,
                    cess, total_tax, doc_type, status, saved_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING', ?)
            """, (
                period_label,
                gstin,
                clean_inv,
                str(row.get("Raw_Invoice_Number", clean_inv)),
                str(row.get("Raw_Invoice_Date", "")) if pd.notna(row.get("Raw_Invoice_Date")) else None,
                str(row.get("Raw_Vendor_Name", "")) if pd.notna(row.get("Raw_Vendor_Name")) else "Unknown Vendor",
                float(row.get("Parsed_Taxable_Value", 0.0)) if pd.notna(row.get("Parsed_Taxable_Value")) else 0.0,
                float(row.get("Parsed_IGST", 0.0)) if pd.notna(row.get("Parsed_IGST")) else 0.0,
                float(row.get("Parsed_CGST", 0.0)) if pd.notna(row.get("Parsed_CGST")) else 0.0,
                float(row.get("Parsed_SGST", 0.0)) if pd.notna(row.get("Parsed_SGST")) else 0.0,
                float(row.get("Parsed_CESS", 0.0)) if pd.notna(row.get("Parsed_CESS")) else 0.0,
                float(row.get("Parsed_Total_Tax", 0.0)) if pd.notna(row.get("Parsed_Total_Tax")) else 0.0,
                str(row.get("Parsed_Document_Type", "INVOICE")),
                now_str
            ))
            if cursor.rowcount > 0:
                saved_count += 1
        except Exception:
            continue

    conn.commit()
    conn.close()
    return saved_count


def get_all_pending_invoices(db_path: str = DEFAULT_DB_PATH) -> pd.DataFrame:
    """Returns all currently pending carried-forward invoices as a DataFrame."""
    init_database(db_path)
    conn = get_connection(db_path)
    df = pd.read_sql_query(
        "SELECT * FROM pending_itc_ledger WHERE status = 'PENDING' ORDER BY period_label, vendor_name",
        conn
    )
    conn.close()
    return df


def resolve_late_filings(
    unmatched_gstr2b_df: pd.DataFrame,
    current_period: str,
    db_path: str = DEFAULT_DB_PATH
) -> pd.DataFrame:
    """
    Checks GSTR-2B invoices that were not matched against current-month books
    against past-month pending invoices stored in the SQLite ledger.
    If a match is found:
      - The database record status is updated to 'RESOLVED'.
      - A record is returned detailing the vendor, invoice, original period,
        and tax recovered.
    """
    if unmatched_gstr2b_df.empty:
        return pd.DataFrame()

    init_database(db_path)
    conn = get_connection(db_path)
    cursor = conn.cursor()
    resolved_records = []
    now_str = datetime.now().isoformat()

    for _, g2b_row in unmatched_gstr2b_df.iterrows():
        g_gstin = str(g2b_row.get("Parsed_GSTIN", "")).strip()
        g_inv = str(g2b_row.get("Parsed_Inv", "")).strip()

        if not g_gstin or g_gstin == "UNKNOWN_GSTIN" or not g_inv or g_inv == "UNKNOWN_INV":
            continue

        # Check if there is a pending invoice in an earlier month
        cursor.execute("""
            SELECT * FROM pending_itc_ledger 
            WHERE gstin = ? AND clean_invoice_number = ? AND status = 'PENDING' AND period_label != ?
            LIMIT 1
        """, (g_gstin, g_inv, current_period))

        match = cursor.fetchone()
        if match:
            match_id = match["id"]
            orig_period = match["period_label"]
            raw_inv = match["raw_invoice_number"]
            vendor = match["vendor_name"]
            tax = match["total_tax"]
            taxable = match["taxable_value"]

            # Mark as resolved
            cursor.execute("""
                UPDATE pending_itc_ledger 
                SET status = 'RESOLVED', resolved_period = ?, resolved_at = ?
                WHERE id = ?
            """, (current_period, now_str, match_id))

            resolved_records.append({
                "GSTIN": g_gstin,
                "Vendor_Name": vendor,
                "Invoice_Number": raw_inv,
                "Original_Period": orig_period,
                "Resolved_In_Period": current_period,
                "Taxable_Value": taxable,
                "Total_Tax": tax,
                "Resolution_Note": f"Late filing resolved: originally missed in {orig_period}, appeared in GSTR-2B of {current_period}"
            })

    conn.commit()
    conn.close()

    if not resolved_records:
        return pd.DataFrame()
    return pd.DataFrame(resolved_records)


def clear_ledger(db_path: str = DEFAULT_DB_PATH) -> None:
    """Clears all ledger records (useful for testing or resetting)."""
    init_database(db_path)
    conn = get_connection(db_path)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM pending_itc_ledger;")
    conn.commit()
    conn.close()
