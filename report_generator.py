"""
report_generator.py - Comprehensive Excel & Communication Report Generator
GSTR2B Match Application
=========================================================================
Generates:
1. Multi-tab, color-coded, audit-ready Excel reports using openpyxl.
2. Formatted vendor follow-up emails and WhatsApp messages for missing ITC.
"""

import io
from datetime import datetime
from typing import Dict, Any, List
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def generate_excel_reconciliation_report(
    books_df: pd.DataFrame,
    missing_in_books_df: pd.DataFrame,
    metrics: Dict[str, Any]
) -> io.BytesIO:
    """
    Generates a beautifully styled, color-coded Excel workbook with multiple tabs:
    - Summary Dashboard (KPIs, trapped ITC, category counts)
    - Exact Matches
    - Fuzzy Matches
    - Mismatches (Tax/Date discrepancy)
    - Missing in GSTR-2B (Trapped ITC)
    - Missing in Purchase Book
    - RCM & Notes
    """
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles
    font_family = "Segoe UI"
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    title_font = Font(name=font_family, size=16, bold=True, color="1E3A8A")
    subtitle_font = Font(name=font_family, size=11, italic=True, color="4B5563")
    kpi_title_font = Font(name=font_family, size=10, bold=True, color="374151")
    kpi_val_font = Font(name=font_family, size=14, bold=True, color="1E3A8A")
    kpi_val_danger_font = Font(name=font_family, size=14, bold=True, color="DC2626")
    kpi_val_success_font = Font(name=font_family, size=14, bold=True, color="16A34A")

    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )

    fill_navy = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_green = PatternFill(start_color="16A34A", end_color="16A34A", fill_type="solid")
    fill_yellow = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")
    fill_red = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid")
    fill_purple = PatternFill(start_color="7C3AED", end_color="7C3AED", fill_type="solid")
    fill_gray = PatternFill(start_color="4B5563", end_color="4B5563", fill_type="solid")
    fill_card = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")

    # -------------------------------------------------------------------------
    # TAB 1: SUMMARY DASHBOARD
    # -------------------------------------------------------------------------
    ws_sum = wb.create_sheet(title="Summary Dashboard")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum["B2"] = "GSTR-2B Input Tax Credit (ITC) Reconciliation Report"
    ws_sum["B2"].font = title_font
    ws_sum["B3"] = f"Generated on {datetime.now().strftime('%d-%b-%Y %H:%M:%S')} | GSTR2B Match Engine"
    ws_sum["B3"].font = subtitle_font

    # KPI Block 1: Total ITC Claimable vs Trapped
    kpis = [
        ("Total Invoices (Books)", f"{metrics.get('total_books_rows', 0):,}", kpi_val_font),
        ("Total ITC Claimed (Matched)", f"₹{metrics.get('total_itc_claimed', 0.0):,.2f}", kpi_val_success_font),
        ("Trapped / At-Risk ITC (Missing 2B)", f"₹{metrics.get('total_itc_blocked_risk', 0.0):,.2f}", kpi_val_danger_font),
        ("Discrepant ITC (Mismatches)", f"₹{metrics.get('total_mismatch_itc', 0.0):,.2f}", kpi_val_font),
        ("Unrecorded ITC (On Portal)", f"₹{metrics.get('unrecorded_2b_itc', 0.0):,.2f}", kpi_val_font),
    ]

    col_idx = 2
    for title, val, vfont in kpis:
        c_title = ws_sum.cell(row=5, column=col_idx, value=title)
        c_title.font = kpi_title_font
        c_title.fill = fill_card
        c_title.alignment = Alignment(horizontal="center", vertical="center")
        c_title.border = thin_border

        c_val = ws_sum.cell(row=6, column=col_idx, value=val)
        c_val.font = vfont
        c_val.fill = fill_card
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.border = thin_border
        
        ws_sum.column_dimensions[get_column_letter(col_idx)].width = 28
        col_idx += 1

    # Breakdown Table
    ws_sum["B9"] = "Category Breakdown & Counts"
    ws_sum["B9"].font = Font(name=font_family, size=13, bold=True, color="1E3A8A")

    breakdown_headers = ["Reconciliation Category", "Invoice Count", "Total Tax Value (₹)", "Status / Action Needed"]
    for j, h in enumerate(breakdown_headers, start=2):
        cell = ws_sum.cell(row=11, column=j, value=h)
        cell.font = header_font
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = thin_border

    breakdown_data = [
        ("Exact Match", metrics.get("exact_matches", 0), metrics.get("total_itc_claimed", 0.0), "Eligible for 100% ITC claim without risk", fill_green),
        ("Fuzzy Match", metrics.get("fuzzy_matches", 0), 0.0, "Matched with invoice typo/format diff; check invoice numbers", fill_yellow),
        ("Mismatch", metrics.get("mismatches", 0), metrics.get("total_mismatch_itc", 0.0), "Rate or date discrepancy; verify bill copy with vendor", fill_yellow),
        ("Missing in GSTR-2B", metrics.get("missing_in_2b", 0), metrics.get("total_itc_blocked_risk", 0.0), "BLOCKS ITC: Vendor has not uploaded invoice on GST portal", fill_red),
        ("Missing in Purchase Book", metrics.get("missing_in_books", 0), metrics.get("unrecorded_2b_itc", 0.0), "Portal shows bill, but books don't; check missing purchase entries", fill_purple),
        ("RCM Invoices (Isolated)", metrics.get("rcm_count", 0), 0.0, "Reverse charge: Tax to be paid in cash, claimed separately", fill_gray)
    ]

    for i, (cat, cnt, amt, desc, col_fill) in enumerate(breakdown_data, start=12):
        c1 = ws_sum.cell(row=i, column=2, value=cat)
        c2 = ws_sum.cell(row=i, column=3, value=cnt)
        c3 = ws_sum.cell(row=i, column=4, value=amt)
        c4 = ws_sum.cell(row=i, column=5, value=desc)
        
        c3.number_format = '#,##0.00'
        for c in (c1, c2, c3, c4):
            c.border = thin_border
            c.font = Font(name=font_family, size=10)

    # -------------------------------------------------------------------------
    # DATA SHEETS
    # -------------------------------------------------------------------------
    sheets_config = [
        ("Exact Matches", books_df[books_df["Match_Category"] == "Exact Match"], fill_green),
        ("Fuzzy Matches", books_df[books_df["Match_Category"] == "Fuzzy Match"], fill_yellow),
        ("Mismatches", books_df[books_df["Match_Category"] == "Mismatch"], fill_yellow),
        ("Missing in GSTR-2B", books_df[books_df["Match_Category"] == "Missing in GSTR-2B"], fill_red),
        ("Missing in Books", missing_in_books_df, fill_purple),
        ("RCM & Others", books_df[books_df["Match_Category"] == "RCM Invoice (Isolated)"], fill_gray),
    ]

    display_cols = [
        ("Raw_GSTIN", "Supplier GSTIN"),
        ("Raw_Vendor_Name", "Supplier Name"),
        ("Raw_Invoice_Number", "Books Inv No"),
        ("Matched_2B_Invoice", "GSTR-2B Inv No"),
        ("Raw_Invoice_Date", "Invoice Date"),
        ("Parsed_Taxable_Value", "Taxable Value (₹)"),
        ("Parsed_IGST", "IGST (₹)"),
        ("Parsed_CGST", "CGST (₹)"),
        ("Parsed_SGST", "SGST (₹)"),
        ("Parsed_CESS", "Cess (₹)"),
        ("Parsed_Total_Tax", "Total Tax (₹)"),
        ("Match_Confidence", "Confidence %"),
        ("Discrepancy_Note", "Discrepancy / Remarks")
    ]

    for sheet_title, sub_df, header_color in sheets_config:
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        # Write Headers
        for col_num, (orig_col, label) in enumerate(display_cols, start=1):
            c = ws.cell(row=1, column=col_num, value=label)
            c.font = header_font
            c.fill = header_color
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border

        # Write Data
        for row_num, (_, r) in enumerate(sub_df.iterrows(), start=2):
            for col_num, (orig_col, _) in enumerate(display_cols, start=1):
                val = r.get(orig_col, "")
                cell = ws.cell(row=row_num, column=col_num, value=val)
                cell.border = thin_border
                cell.font = Font(name=font_family, size=10)

                # Format numbers
                if "Value" in orig_col or "Tax" in orig_col or orig_col in ["Parsed_IGST", "Parsed_CGST", "Parsed_SGST", "Parsed_CESS"]:
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00'

        # Auto-adjust column widths
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    output.seek(0)
    return output


def generate_vendor_email(
    vendor_name: str,
    gstin: str,
    invoices_list: List[Dict[str, Any]],
    recipient_email: str = ""
) -> str:
    """
    Generates a courteous, professional vendor follow-up email in plain language.
    Lists pending invoices, dates, taxable values, and tax amounts.
    """
    total_blocked_tax = sum(inv.get("total_tax", 0.0) for inv in invoices_list)
    inv_count = len(invoices_list)

    table_rows = []
    for inv in invoices_list:
        table_rows.append(
            f"  • Inv #{inv.get('invoice_no', 'N/A')} dated {inv.get('date', 'N/A')} "
            f"| Taxable: ₹{inv.get('taxable_value', 0.0):,.2f} "
            f"| GST: ₹{inv.get('total_tax', 0.0):,.2f}"
        )
    inv_lines = "\n".join(table_rows)

    email_body = f"""Subject: Urgent: GSTR-1 Filing Discrepancy / Missing Invoices for Input Tax Credit (GSTIN: {gstin})

Dear Accounts Team ({vendor_name}),

Greetings!

While reconciling our purchase register with our government GSTR-2B report, we noticed that {inv_count} invoice(s) issued by your company are currently not appearing on the GST portal:

{inv_lines}

Total Input Tax Credit (ITC) Blocked: ₹{total_blocked_tax:,.2f}

As per Indian GST regulations, we cannot claim Input Tax Credit on these invoices until they are successfully uploaded in your GSTR-1 / IFF and reflected in our GSTR-2B.

Could you please verify the following:
1. Were these invoices filed under the correct GSTIN ({gstin})?
2. Were they reported in your latest GSTR-1 return, or will they be included in the upcoming filing?
3. If there was any invoice number typo or amendment, kindly share the corrected details.

We request you to kindly check this at your earliest convenience to avoid interest and credit reversal on our end.

Thank you for your prompt assistance!

Best regards,
Accounts & Finance Department
"""
    return email_body


def generate_vendor_whatsapp(
    vendor_name: str,
    gstin: str,
    inv_count: int,
    total_tax: float
) -> str:
    """Generates a concise WhatsApp follow-up message."""
    return (
        f"Hi {vendor_name}, regarding our GST reconciliation: {inv_count} invoice(s) "
        f"under GSTIN {gstin} are missing from our GSTR-2B portal report (blocked ITC ₹{total_tax:,.2f}). "
        f"Could you please check your GSTR-1 filing status so our tax credit is not blocked? Thank you!"
    )
