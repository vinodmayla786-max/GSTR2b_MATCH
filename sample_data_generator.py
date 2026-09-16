"""
sample_data_generator.py - Synthetic Test Datasets for GST Reconciliation
GSTR2B Match Application
========================================================================
Generates realistic sample Purchase Book and GSTR-2B datasets for immediate
zero-data testing. Includes exact matches, typos, tax mismatches, missing
vendor filings, and RCM bills.
"""

import pandas as pd


def get_sample_purchase_book() -> pd.DataFrame:
    """Returns a realistic synthetic Purchase Book DataFrame."""
    data = [
        {
            "Supplier GSTIN": "27AABCU9603R1ZM",
            "Supplier Name": "Tata Steel Limited",
            "Invoice Number": "TSL/2026/101",
            "Invoice Date": "05-08-2026",
            "Taxable Value": 100000.0,
            "IGST": 18000.0,
            "CGST": 0.0,
            "SGST": 0.0,
            "CESS": 0.0,
            "Total Tax": 18000.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        {
            "Supplier GSTIN": "07AAACG0532F1Z8",
            "Supplier Name": "Godrej Consumer Products",
            "Invoice Number": "GCPL-0452",
            "Invoice Date": "10-08-2026",
            "Taxable Value": 50000.0,
            "IGST": 0.0,
            "CGST": 4500.0,
            "SGST": 4500.0,
            "CESS": 0.0,
            "Total Tax": 9000.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        {
            "Supplier GSTIN": "29AABCB1234F1Z5",
            "Supplier Name": "Infosys BPM Tech Supplies",
            "Invoice Number": "INF-2026-889",
            "Invoice Date": "12-08-2026",
            "Taxable Value": 75000.0,
            "IGST": 13500.0,
            "CGST": 0.0,
            "SGST": 0.0,
            "CESS": 0.0,
            "Total Tax": 13500.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        {
            "Supplier GSTIN": "06AABCI9876Q1Z2",
            "Supplier Name": "Apex Logistics Private Ltd",
            "Invoice Number": "APX/AUG/012",
            "Invoice Date": "15-08-2026",
            "Taxable Value": 120000.0,
            "IGST": 21600.0,
            "CGST": 0.0,
            "SGST": 0.0,
            "CESS": 0.0,
            "Total Tax": 21600.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        {
            "Supplier GSTIN": "33AAACB9999P1Z1",
            "Supplier Name": "Kaveri Engineering Works",
            "Invoice Number": "KEW-9901",
            "Invoice Date": "18-08-2026",
            "Taxable Value": 250000.0,
            "IGST": 45000.0,
            "CGST": 0.0,
            "SGST": 0.0,
            "CESS": 0.0,
            "Total Tax": 45000.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        {
            "Supplier GSTIN": "24AAACG1111B1Z4",
            "Supplier Name": "Shree Ram Packaging",
            "Invoice Number": "SRP-0077",
            "Invoice Date": "22-08-2026",
            "Taxable Value": 30000.0,
            "IGST": 0.0,
            "CGST": 2700.0,
            "SGST": 2700.0,
            "CESS": 0.0,
            "Total Tax": 5400.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        {
            "Supplier GSTIN": "27AABCZ5555M1Z7",
            "Supplier Name": "Suraksha Security Services",
            "Invoice Number": "SEC-AUG-2026",
            "Invoice Date": "28-08-2026",
            "Taxable Value": 40000.0,
            "IGST": 0.0,
            "CGST": 3600.0,
            "SGST": 3600.0,
            "CESS": 0.0,
            "Total Tax": 7200.0,
            "Reverse Charge": "Y",
            "Document Type": "INVOICE"
        }
    ]
    return pd.DataFrame(data)


def get_sample_gstr2b() -> pd.DataFrame:
    """Returns a realistic synthetic GSTR-2B DataFrame reflecting portal real-world quirks."""
    data = [
        # Exact Match with Tata Steel
        {
            "Supplier GSTIN": "27AABCU9603R1ZM",
            "Supplier Name": "Tata Steel Limited",
            "Invoice Number": "TSL/2026/101",
            "Invoice Date": "05-08-2026",
            "Taxable Value": 100000.0,
            "IGST": 18000.0,
            "CGST": 0.0,
            "SGST": 0.0,
            "CESS": 0.0,
            "Total Tax": 18000.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        # Fuzzy Match with Godrej (Vendor entered without hyphen: GCPL0452)
        {
            "Supplier GSTIN": "07AAACG0532F1Z8",
            "Supplier Name": "Godrej Consumer Products Ltd",
            "Invoice Number": "GCPL0452",
            "Invoice Date": "10-08-2026",
            "Taxable Value": 50000.0,
            "IGST": 0.0,
            "CGST": 4500.0,
            "SGST": 4500.0,
            "CESS": 0.0,
            "Total Tax": 9000.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        # Mismatch with Infosys (Vendor filed higher amount: 80,000 taxable vs 75,000 in books)
        {
            "Supplier GSTIN": "29AABCB1234F1Z5",
            "Supplier Name": "Infosys BPM Tech Supplies",
            "Invoice Number": "INF-2026-889",
            "Invoice Date": "12-08-2026",
            "Taxable Value": 80000.0,
            "IGST": 14400.0,
            "CGST": 0.0,
            "SGST": 0.0,
            "CESS": 0.0,
            "Total Tax": 14400.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        # Missing in 2B: Kaveri Engineering Works (KEW-9901) is NOT uploaded by vendor! (Blocks 45,000 ITC)
        
        # Missing in Purchase Book: Vendor filed an invoice not recorded internally yet
        {
            "Supplier GSTIN": "06AABCI9876Q1Z2",
            "Supplier Name": "Apex Logistics Private Ltd",
            "Invoice Number": "APX/AUG/099",
            "Invoice Date": "25-08-2026",
            "Taxable Value": 35000.0,
            "IGST": 6300.0,
            "CGST": 0.0,
            "SGST": 0.0,
            "CESS": 0.0,
            "Total Tax": 6300.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        # Shree Ram Packaging: Matched with slight date variation (21-08-2026 vs 22-08-2026)
        {
            "Supplier GSTIN": "24AAACG1111B1Z4",
            "Supplier Name": "Shree Ram Packaging",
            "Invoice Number": "SRP-0077",
            "Invoice Date": "21-08-2026",
            "Taxable Value": 30000.0,
            "IGST": 0.0,
            "CGST": 2700.0,
            "SGST": 2700.0,
            "CESS": 0.0,
            "Total Tax": 5400.0,
            "Reverse Charge": "N",
            "Document Type": "INVOICE"
        },
        # Suraksha Security RCM
        {
            "Supplier GSTIN": "27AABCZ5555M1Z7",
            "Supplier Name": "Suraksha Security Services",
            "Invoice Number": "SEC-AUG-2026",
            "Invoice Date": "28-08-2026",
            "Taxable Value": 40000.0,
            "IGST": 0.0,
            "CGST": 3600.0,
            "SGST": 3600.0,
            "CESS": 0.0,
            "Total Tax": 7200.0,
            "Reverse Charge": "Y",
            "Document Type": "INVOICE"
        }
    ]
    return pd.DataFrame(data)
