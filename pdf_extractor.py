"""
pdf_extractor.py - Unstructured GST PDF to Structured Data & Tally XML
GSTR2B Match Application
=====================================================================
Extracts structured table data from messy, multi-page, merged-cell GST PDFs
(GSTR-2B, GSTR-1, GSTR-3B) and generates clean CSV/Excel & Tally XML.
"""

import re
import io
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Tuple, List, Dict, Any, Optional
import pandas as pd
import numpy as np
import pdfplumber


KNOWN_COLUMN_SYNONYMS = {
    "GSTIN": [
        "gstin", "gstin of supplier", "supplier gstin", "gstin/uin",
        "gstin / uin", "receiver gstin", "recipient gstin", "party gstin"
    ],
    "Vendor_Name": [
        "trade/legal name", "legal name", "trade name", "supplier name",
        "party name", "name of supplier", "name of party", "vendor name"
    ],
    "Invoice_Number": [
        "invoice number", "invoice no", "invoice no.", "inv no", "inv no.",
        "document number", "doc no", "doc no.", "note number", "bill no"
    ],
    "Invoice_Type": [
        "invoice type", "doc type", "document type", "type", "note type"
    ],
    "Invoice_Date": [
        "invoice date", "inv date", "document date", "doc date", "date", "bill date"
    ],
    "Invoice_Value": [
        "invoice value", "inv value", "total value", "invoice amount", "doc value"
    ],
    "Place_Of_Supply": [
        "place of supply", "pos", "state code", "place of supply (pos)"
    ],
    "Reverse_Charge": [
        "supply attract reverse charge", "reverse charge", "rcm", "rev charge", "applicable %"
    ],
    "Tax_Rate": [
        "rate", "tax rate", "rate (%)", "tax rate (%)"
    ],
    "Taxable_Value": [
        "taxable value", "taxable value (₹)", "taxable value(₹)", "taxable amt", "taxable amount"
    ],
    "IGST": [
        "integrated tax", "integrated tax (₹)", "igst", "igst amount", "igst (₹)"
    ],
    "CGST": [
        "central tax", "central tax (₹)", "cgst", "cgst amount", "cgst (₹)"
    ],
    "SGST": [
        "state/ut tax", "state/ut tax (₹)", "state tax", "sgst", "sgst/utgst", "sgst amount"
    ],
    "CESS": [
        "cess", "cess (₹)", "cess amount"
    ]
}


def clean_numeric(val: Any) -> float:
    """
    Cleans Indian number formats (e.g. 1,25,000.50, ₹ 500, (100.00)).
    Returns 0.0 if cannot be parsed.
    """
    if val is None or pd.isna(val):
        return 0.0
    s = str(val).strip()
    if not s or s.lower() in ["nil", "none", "-", "na", "n/a"]:
        return 0.0
    
    # Handle negative numbers in parentheses (e.g. Credit Notes: (1,250.00))
    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()

    # Remove currency symbols, commas, and spaces
    s = re.sub(r"[₹Rs\.,\s]", lambda m: "." if m.group(0) == "." else "", s)
    # Fix multiple periods if any
    parts = s.split(".")
    if len(parts) > 2:
        s = "".join(parts[:-1]) + "." + parts[-1]

    try:
        num = float(s)
        return -num if is_negative else num
    except ValueError:
        return 0.0


def standardize_date(date_str: Any) -> str:
    """Standardizes various date formats into DD-MM-YYYY."""
    if not date_str or pd.isna(date_str):
        return ""
    s = str(date_str).strip()
    # Common formats in Indian GST returns: DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d-%b-%y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            dt = datetime.strptime(s, fmt)
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            continue
    return s


def identify_table_headers(header_row: List[Any]) -> Dict[int, str]:
    """
    Matches raw header texts against known Indian GST column synonyms.
    Returns a mapping of column_index -> Standardized_Field_Name.
    """
    col_mapping = {}
    used_standards = set()

    for idx, cell in enumerate(header_row):
        if not cell:
            continue
        cleaned = re.sub(r"\s+", " ", str(cell).lower().strip())
        
        # Check against known synonyms
        matched_standard = None
        for standard_name, synonyms in KNOWN_COLUMN_SYNONYMS.items():
            if standard_name in used_standards:
                continue
            for syn in synonyms:
                if syn == cleaned or (len(syn) > 4 and syn in cleaned):
                    matched_standard = standard_name
                    break
            if matched_standard:
                break
        
        if matched_standard:
            col_mapping[idx] = matched_standard
            used_standards.add(matched_standard)

    return col_mapping


def extract_tables_from_pdf(pdf_file_bytes) -> Tuple[pd.DataFrame, List[str], Dict[str, Any]]:
    """
    Extracts structured invoice rows from a PDF file using pdfplumber.
    Handles:
    - Multi-page tables
    - Merged cells across rows (forward filling supplier info)
    - Multiline headers
    - Credit/Debit note indicators
    """
    warnings = []
    stats = {"pages_scanned": 0, "tables_found": 0, "rows_extracted": 0}
    all_rows = []
    
    try:
        with pdfplumber.open(pdf_file_bytes) as pdf:
            stats["pages_scanned"] = len(pdf.pages)
            active_col_mapping = {}
            last_seen_gstin = ""
            last_seen_vendor = ""

            for page_idx, page in enumerate(pdf.pages):
                tables = page.extract_tables()
                if not tables:
                    continue

                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    stats["tables_found"] += 1

                    # Check if first row is a header
                    row_0 = table[0]
                    detected_mapping = identify_table_headers(row_0)

                    # Sometimes header is split across rows 0 and 1 (nested/merged headers)
                    if len(detected_mapping) < 3 and len(table) > 2:
                        combined_header = [
                            f"{table[0][i] or ''} {table[1][i] or ''}".strip()
                            for i in range(min(len(table[0]), len(table[1])))
                        ]
                        nested_mapping = identify_table_headers(combined_header)
                        if len(nested_mapping) > len(detected_mapping):
                            detected_mapping = nested_mapping
                            data_start_idx = 2
                        else:
                            data_start_idx = 1
                    else:
                        data_start_idx = 1

                    # Update mapping if this table has strong headers
                    if len(detected_mapping) >= 3:
                        active_col_mapping = detected_mapping
                    elif not active_col_mapping:
                        # Skip table if no headers ever discovered
                        continue

                    # Process data rows
                    for r_idx in range(data_start_idx, len(table)):
                        row = table[r_idx]
                        if not row or all(c is None or str(c).strip() == "" for c in row):
                            continue

                        parsed_record = {}
                        for col_idx, std_name in active_col_mapping.items():
                            if col_idx < len(row):
                                val = row[col_idx]
                                parsed_record[std_name] = str(val).strip() if val is not None else ""

                        # Check for merged cell / continuation:
                        # If GSTIN or Vendor is missing, but invoice number is present, carry forward last seen
                        curr_gstin = parsed_record.get("GSTIN", "").replace(" ", "").upper()
                        curr_vendor = parsed_record.get("Vendor_Name", "").strip()
                        curr_inv = parsed_record.get("Invoice_Number", "").strip()

                        if curr_gstin and len(curr_gstin) >= 15:
                            last_seen_gstin = curr_gstin
                        elif not curr_gstin and last_seen_gstin and curr_inv:
                            parsed_record["GSTIN"] = last_seen_gstin

                        if curr_vendor:
                            last_seen_vendor = curr_vendor
                        elif not curr_vendor and last_seen_vendor and curr_inv:
                            parsed_record["Vendor_Name"] = last_seen_vendor

                        # Filter out subtotal / total rows
                        inv_val_str = parsed_record.get("Invoice_Number", "").lower()
                        if any(kw in inv_val_str for kw in ["total", "subtotal", "grand total", "page total"]):
                            continue

                        # If invoice number or GSTIN is present, record it
                        if parsed_record.get("Invoice_Number") or parsed_record.get("Taxable_Value"):
                            parsed_record["Page_Number"] = page_idx + 1
                            all_rows.append(parsed_record)

    except Exception as e:
        warnings.append(f"PDF extraction error: {str(e)}")

    if not all_rows:
        warnings.append("No structured GST table data could be detected. Please ensure this is an official GSTR-2B, GSTR-1, or GSTR-3B PDF.")
        return pd.DataFrame(), warnings, stats

    df = pd.DataFrame(all_rows)

    # Standardize column values
    if "GSTIN" in df.columns:
        df["GSTIN"] = df["GSTIN"].astype(str).str.replace(r"\s+", "", regex=True).str.upper()
    else:
        df["GSTIN"] = ""

    if "Invoice_Number" in df.columns:
        df["Invoice_Number"] = df["Invoice_Number"].astype(str).str.strip()
    else:
        df["Invoice_Number"] = ""

    if "Invoice_Date" in df.columns:
        df["Invoice_Date"] = df["Invoice_Date"].apply(standardize_date)
    else:
        df["Invoice_Date"] = ""

    # Clean numeric columns
    numeric_targets = ["Invoice_Value", "Taxable_Value", "IGST", "CGST", "SGST", "CESS"]
    for col in numeric_targets:
        if col in df.columns:
            df[col] = df[col].apply(clean_numeric)
        else:
            df[col] = 0.0

    # Calculate Total_Tax if not already given
    df["Total_Tax"] = df["IGST"] + df["CGST"] + df["SGST"] + df["CESS"]

    # Calculate Total Invoice Value if missing or 0
    df["Invoice_Value"] = np.where(
        df["Invoice_Value"] > 0,
        df["Invoice_Value"],
        df["Taxable_Value"] + df["Total_Tax"]
    )

    stats["rows_extracted"] = len(df)
    return df, warnings, stats


def generate_tally_xml(df: pd.DataFrame, company_name: str = "Company Name") -> str:
    """
    Generates Tally-compatible XML (Vouchers import format) from extracted GST records.
    Can be imported into TallyPrime or Tally.ERP 9 via Gateway of Tally -> Import Data -> Vouchers.
    """
    envelope = ET.Element("ENVELOPE")
    
    header = ET.SubElement(envelope, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"
    
    body = ET.SubElement(envelope, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    
    req_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(req_desc, "REPORTNAME").text = "Vouchers"
    
    static_vars = ET.SubElement(req_desc, "STATICVARIABLES")
    ET.SubElement(static_vars, "SVCURRENTCOMPANY").text = company_name
    
    req_data = ET.SubElement(import_data, "REQUESTDATA")
    
    for idx, row in df.iterrows():
        tally_msg = ET.SubElement(req_data, "TALLYMESSAGE")
        tally_msg.set("xmlns:UDF", "TallyUDF")
        
        # Voucher node
        voucher = ET.SubElement(tally_msg, "VOUCHER")
        vch_type = "Purchase"
        voucher.set("VCHTYPE", vch_type)
        voucher.set("ACTION", "Create")
        
        # Date in YYYYMMDD format
        date_str = str(row.get("Invoice_Date", ""))
        try:
            dt = datetime.strptime(date_str, "%d-%m-%Y")
            tally_date = dt.strftime("%Y%m%d")
            display_date = dt.strftime("%d-%b-%Y")
        except Exception:
            tally_date = datetime.now().strftime("%Y%m%d")
            display_date = datetime.now().strftime("%d-%b-%Y")
            
        ET.SubElement(voucher, "DATE").text = tally_date
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = vch_type
        
        inv_no = str(row.get("Invoice_Number", f"INV-{idx+1}"))
        ET.SubElement(voucher, "REFERENCE").text = inv_no
        ET.SubElement(voucher, "VOUCHERNUMBER").text = inv_no
        
        party_name = str(row.get("Vendor_Name", "")).strip() or f"Supplier {row.get('GSTIN', '')}"
        ET.SubElement(voucher, "PARTYLEDGERNAME").text = party_name
        
        narration = f"GST Auto-extracted: Inv #{inv_no} dated {display_date} from {party_name} (GSTIN: {row.get('GSTIN', '')})"
        ET.SubElement(voucher, "NARRATION").text = narration
        
        total_inv_val = float(row.get("Invoice_Value", 0.0))
        taxable_val = float(row.get("Taxable_Value", 0.0))
        igst_val = float(row.get("IGST", 0.0))
        cgst_val = float(row.get("CGST", 0.0))
        sgst_val = float(row.get("SGST", 0.0))
        cess_val = float(row.get("CESS", 0.0))

        # 1. Credit Party Ledger (Total Invoice Value)
        ledger_party = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(ledger_party, "LEDGERNAME").text = party_name
        ET.SubElement(ledger_party, "ISDEEMEDPOSITIVE").text = "No"  # Credit
        ET.SubElement(ledger_party, "AMOUNT").text = f"{total_inv_val:.2f}"
        
        # 2. Debit Purchase Account (Taxable Value)
        if taxable_val > 0:
            ledger_purch = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(ledger_purch, "LEDGERNAME").text = "GST Purchase"
            ET.SubElement(ledger_purch, "ISDEEMEDPOSITIVE").text = "Yes"  # Debit
            ET.SubElement(ledger_purch, "AMOUNT").text = f"-{taxable_val:.2f}"
            
        # 3. Debit Tax Ledgers
        if igst_val > 0:
            l_igst = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l_igst, "LEDGERNAME").text = "Input IGST"
            ET.SubElement(l_igst, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(l_igst, "AMOUNT").text = f"-{igst_val:.2f}"
            
        if cgst_val > 0:
            l_cgst = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l_cgst, "LEDGERNAME").text = "Input CGST"
            ET.SubElement(l_cgst, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(l_cgst, "AMOUNT").text = f"-{cgst_val:.2f}"
            
        if sgst_val > 0:
            l_sgst = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l_sgst, "LEDGERNAME").text = "Input SGST"
            ET.SubElement(l_sgst, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(l_sgst, "AMOUNT").text = f"-{sgst_val:.2f}"

        if cess_val > 0:
            l_cess = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(l_cess, "LEDGERNAME").text = "Input Cess"
            ET.SubElement(l_cess, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(l_cess, "AMOUNT").text = f"-{cess_val:.2f}"

    # Return XML as string
    xml_declaration = '<?xml version="1.0" encoding="utf-8"?>\n'
    return xml_declaration + ET.tostring(envelope, encoding="utf-8").decode("utf-8")
