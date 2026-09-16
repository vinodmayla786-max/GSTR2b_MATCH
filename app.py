"""
app.py - Main Streamlit Web Application for GSTR2B Match
Indian Tax Compliance & Input Tax Credit (ITC) Reconciliation Engine
===================================================================
Production-ready application designed for MSMEs, accountants, and CA firms.
Includes:
1. PDF Extractor (Messy GST PDF -> Structured Excel/CSV & Tally XML)
2. Reconciliation Engine (Purchase Book vs GSTR-2B with fuzzy matching & tolerances)
3. Reports & Dashboard (Visual analytics, styled Excel download, vendor follow-up emails)
"""

import io
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# Local helper modules
import db
import matcher
import pdf_extractor
import report_generator
import sample_data_generator

# Streamlit Page Setup
st.set_page_config(
    page_title="GSTR2B Match - Indian GST ITC Reconciler",
    page_icon="🇮🇳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling for polished, professional typography & cards
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #4B5563;
        margin-bottom: 1.5rem;
    }
    .metric-card {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 16px;
        text-align: center;
    }
    .metric-label {
        font-size: 0.85rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
    }
    .metric-val {
        font-size: 1.6rem;
        font-weight: 700;
        color: #0F172A;
    }
    .metric-val-danger {
        color: #DC2626;
    }
    .metric-val-success {
        color: #16A34A;
    }
    .stAlert {
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize Session State
if "books_df" not in st.session_state:
    st.session_state["books_df"] = None
if "gstr2b_df" not in st.session_state:
    st.session_state["gstr2b_df"] = None
if "extracted_pdf_df" not in st.session_state:
    st.session_state["extracted_pdf_df"] = None
if "reco_results" not in st.session_state:
    st.session_state["reco_results"] = None
if "missing_in_books" not in st.session_state:
    st.session_state["missing_in_books"] = None
if "reco_metrics" not in st.session_state:
    st.session_state["reco_metrics"] = None
if "sample_mode_loaded" not in st.session_state:
    st.session_state["sample_mode_loaded"] = False

# Initialize SQLite database
db.init_database()

# =============================================================================
# SIDEBAR NAVIGATION & CONFIGURATION
# =============================================================================
st.sidebar.image("https://raw.githubusercontent.com/FortAwesome/Font-Awesome/6.x/svgs/solid/file-invoice-dollar.svg", width=48)
st.sidebar.markdown("## **GSTR2B Match**")
st.sidebar.caption("Indian GST Reconciliation & Tally Automation")

nav_choice = st.sidebar.radio(
    "Select Workflow Module:",
    ["1. PDF Extractor", "2. Reconciliation", "3. Reports & Dashboard"],
    index=1
)

st.sidebar.divider()

# Sample Data Quick Test Feature
st.sidebar.subheader("🚀 Sample Data Mode")
st.sidebar.caption("Test the app immediately without needing real business files.")
if st.sidebar.button("Load Synthetic Test Data", use_container_width=True):
    st.session_state["books_df"] = sample_data_generator.get_sample_purchase_book()
    st.session_state["gstr2b_df"] = sample_data_generator.get_sample_gstr2b()
    st.session_state["sample_mode_loaded"] = True
    
    # Auto-run reconciliation with default parameters
    books, missing_b, all_2b, metrics = matcher.perform_reconciliation(
        st.session_state["books_df"],
        st.session_state["gstr2b_df"]
    )
    st.session_state["reco_results"] = books
    st.session_state["missing_in_books"] = missing_b
    st.session_state["reco_metrics"] = metrics
    st.sidebar.success("✅ Sample files loaded & reconciled!")
    st.rerun()

if st.session_state.get("sample_mode_loaded"):
    if st.sidebar.button("Clear Loaded Data", use_container_width=True):
        st.session_state["books_df"] = None
        st.session_state["gstr2b_df"] = None
        st.session_state["reco_results"] = None
        st.session_state["reco_metrics"] = None
        st.session_state["sample_mode_loaded"] = False
        st.rerun()

st.sidebar.divider()
st.sidebar.markdown("### ⚙️ Matching Tolerances")
abs_tolerance = st.sidebar.number_input("Amount Tolerance (₹)", min_value=0.0, max_value=50.0, value=1.0, step=0.5, help="Amounts differing by up to this value (e.g. rounding diff) are considered matching.")
rel_tolerance_pct = st.sidebar.number_input("Amount Tolerance (%)", min_value=0.0, max_value=5.0, value=1.0, step=0.1, help="Percentage difference tolerance.")
date_tolerance_days = st.sidebar.number_input("Date Tolerance (Days)", min_value=0, max_value=30, value=3, step=1, help="Invoice date variation tolerance between books and vendor portal filing.")
min_fuzzy_inv_score = st.sidebar.slider("Fuzzy Invoice Match Threshold (%)", min_value=50, max_value=100, value=75, help="Fuzzy string similarity threshold to catch typographical errors in invoice numbers.")


# =============================================================================
# MODULE 1: UNSTRUCTURED GST PDF -> TALLY EXTRACTOR
# =============================================================================
if nav_choice == "1. PDF Extractor":
    st.markdown('<div class="main-title">📄 Unstructured GST PDF → Tally Extractor</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Extract structured data from merged-cell, multi-page GSTR-2B, GSTR-1, or GSTR-3B PDF files.</div>', unsafe_allow_html=True)

    with st.expander("💡 How this works", expanded=True):
        st.markdown("""
        - **The Problem:** Government GST portal PDF downloads often have merged header cells, page breaks, and missing columns that break normal Excel converters.
        - **The Solution:** This extractor uses table boundary recognition with smart forward-filling of supplier names and GSTINs to generate clean row-by-row data.
        - **Tally Ready:** Download clean Excel/CSV, or export a ready-to-import **Tally XML Voucher** file for instant accounting entries!
        """)

    uploaded_pdf = st.file_uploader(
        "Upload GST PDF (GSTR-2B / GSTR-1 / GSTR-3B)",
        type=["pdf"],
        help="Upload any official government GST return PDF."
    )

    company_for_tally = st.text_input("Tally Company Name (for XML import)", value="My Business Pvt Ltd")

    if uploaded_pdf is not None:
        with st.spinner("Extracting and restructuring table data from PDF..."):
            extracted_df, warnings, stats = pdf_extractor.extract_tables_from_pdf(uploaded_pdf)

        for w in warnings:
            st.warning(f"⚠️ {w}")

        if not extracted_df.empty:
            st.session_state["extracted_pdf_df"] = extracted_df
            st.success(f"✅ Extracted **{stats['rows_extracted']} rows** across **{stats['pages_scanned']} page(s)**!")

            # Metric Bar
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Pages Scanned", stats["pages_scanned"])
            m2.metric("Invoices Extracted", len(extracted_df))
            m3.metric("Total Taxable Value", f"₹{extracted_df['Taxable_Value'].sum():,.2f}")
            m4.metric("Total Tax (IGST+CGST+SGST)", f"₹{extracted_df['Total_Tax'].sum():,.2f}")

            # Preview
            st.markdown("### Extracted Invoice Data Preview")
            st.dataframe(extracted_df, use_container_width=True, height=350)

            # Export actions
            c1, c2, c3, c4 = st.columns(4)

            # Excel download
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                extracted_df.to_excel(writer, index=False, sheet_name="Clean_GST_Data")
            excel_data = excel_buffer.getvalue()

            c1.download_button(
                label="📥 Download Excel (.xlsx)",
                data=excel_data,
                file_name=f"Clean_GST_Extracted_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # CSV download
            csv_data = extracted_df.to_csv(index=False).encode('utf-8')
            c2.download_button(
                label="📥 Download CSV (.csv)",
                data=csv_data,
                file_name=f"Clean_GST_Extracted_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )

            # Tally XML download
            tally_xml_str = pdf_extractor.generate_tally_xml(extracted_df, company_for_tally)
            c3.download_button(
                label="📥 Download Tally XML",
                data=tally_xml_str,
                file_name=f"Tally_Import_Vouchers_{datetime.now().strftime('%Y%m%d')}.xml",
                mime="application/xml",
                use_container_width=True,
                help="Import directly into Tally via Gateway of Tally -> Import Data -> Vouchers."
            )

            # Direct feed to Reconciliation
            if c4.button("⚡ Send to Reconciliation Engine", use_container_width=True, type="primary"):
                st.session_state["gstr2b_df"] = extracted_df
                st.success("Data transferred to Reconciliation! Open tab '2. Reconciliation' above to continue.")


# =============================================================================
# MODULE 2: RECONCILIATION ENGINE
# =============================================================================
elif nav_choice == "2. Reconciliation":
    st.markdown('<div class="main-title">⚖️ Input Tax Credit (ITC) Reconciliation Engine</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Reconcile internal Purchase Books against Government GSTR-2B data with fuzzy matching.</div>', unsafe_allow_html=True)

    with st.expander("💡 How this works", expanded=False):
        st.markdown("""
        - **1. Upload Files:** Upload your internal Purchase Book (.csv or .xlsx) and your GSTR-2B data (.json, .xlsx, .csv, or .pdf).
        - **2. Smart Matching:** The system standardizes invoice numbers (strips slashes, leading zeros, spaces) and matches using GSTIN, invoice date, and amounts within your configured tolerances.
        - **3. Bucketing:** Each invoice is placed into one of 5 categories: Exact Match, Fuzzy Match (typos), Mismatch (value differences), Missing in GSTR-2B (blocked ITC), or Missing in Books.
        - **4. Cross-Month Persistence:** Invoices marked 'Missing in 2B' can be saved to a local SQLite ledger to automatically match when vendors file late in future months!
        """)

    col_b, col_2b = st.columns(2)

    with col_b:
        st.markdown("#### 1. Internal Purchase Book")
        file_books = st.file_uploader("Upload Purchase Book (.csv or .xlsx)", type=["csv", "xlsx"], key="up_books")
        if file_books is not None:
            try:
                if file_books.name.lower().endswith(".csv"):
                    st.session_state["books_df"] = pd.read_csv(file_books)
                else:
                    st.session_state["books_df"] = pd.read_excel(file_books)
                st.success(f"Loaded {len(st.session_state['books_df'])} book records from `{file_books.name}`")
            except Exception as e:
                st.error(f"Error reading Purchase Book: {str(e)}")

        if st.session_state["books_df"] is not None:
            st.caption(f"Active Purchase Book: {len(st.session_state['books_df'])} rows loaded")

    with col_2b:
        st.markdown("#### 2. GSTR-2B Portal Data")
        file_2b = st.file_uploader("Upload GSTR-2B (.json, .xlsx, .csv, .pdf)", type=["json", "xlsx", "csv", "pdf"], key="up_2b")
        if file_2b is not None:
            try:
                fn = file_2b.name.lower()
                if fn.endswith(".json"):
                    g_df, warns = matcher.parse_gstr2b_json(file_2b.read())
                    st.session_state["gstr2b_df"] = g_df
                    for w in warns:
                        st.warning(w)
                elif fn.endswith(".pdf"):
                    with st.spinner("Extracting GSTR-2B from PDF..."):
                        g_df, warns, _ = pdf_extractor.extract_tables_from_pdf(file_2b)
                        st.session_state["gstr2b_df"] = g_df
                        for w in warns:
                            st.warning(w)
                elif fn.endswith(".csv"):
                    st.session_state["gstr2b_df"] = pd.read_csv(file_2b)
                else:
                    st.session_state["gstr2b_df"] = pd.read_excel(file_2b)

                st.success(f"Loaded {len(st.session_state['gstr2b_df'])} GSTR-2B records from `{file_2b.name}`")
            except Exception as e:
                st.error(f"Error reading GSTR-2B: {str(e)}")

        if st.session_state["gstr2b_df"] is not None:
            st.caption(f"Active GSTR-2B: {len(st.session_state['gstr2b_df'])} rows loaded")

    # Filing Period for Cross-Month Carry-Forward
    c_p1, c_p2 = st.columns([2, 1])
    with c_p1:
        current_period = st.text_input("Tax Period Label (e.g. Aug-2026)", value=datetime.now().strftime("%b-%Y"))
    with c_p2:
        check_carry_forward = st.checkbox("Enable Carry-Forward Check", value=True, help="Automatically check if leftover GSTR-2B rows match invoices missed in earlier months.")

    # Run Reconciliation Button
    if st.button("▶ Run ITC Reconciliation", type="primary", use_container_width=True):
        if st.session_state["books_df"] is None or st.session_state["gstr2b_df"] is None:
            st.error("Please upload both your Purchase Book and GSTR-2B data (or click 'Load Synthetic Test Data' in the sidebar).")
        else:
            with st.spinner("Reconciling GSTINs, Invoices, Tax Values, and Tolerances..."):
                books, missing_b, all_2b, metrics = matcher.perform_reconciliation(
                    st.session_state["books_df"],
                    st.session_state["gstr2b_df"],
                    abs_tolerance=abs_tolerance,
                    rel_tolerance_pct=rel_tolerance_pct,
                    date_tolerance_days=date_tolerance_days,
                    min_fuzzy_inv_score=min_fuzzy_inv_score
                )

                # Check Cross-month Carry-Forward in SQLite
                resolved_carry_forward = pd.DataFrame()
                if check_carry_forward:
                    resolved_carry_forward = db.resolve_late_filings(missing_b, current_period)

                st.session_state["reco_results"] = books
                st.session_state["missing_in_books"] = missing_b
                st.session_state["reco_metrics"] = metrics
                st.session_state["resolved_carry_forward"] = resolved_carry_forward
                st.success("Reconciliation completed successfully!")

    # Display Results if Available
    if st.session_state.get("reco_results") is not None:
        books = st.session_state["reco_results"]
        missing_b = st.session_state["missing_in_books"]
        metrics = st.session_state["reco_metrics"]
        resolved_cf = st.session_state.get("resolved_carry_forward", pd.DataFrame())

        st.divider()
        st.markdown("### 📊 Reconciliation Results Overview")

        # Top Metric Cards
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.markdown(f'<div class="metric-card"><div class="metric-label">Total Books Rows</div><div class="metric-val">{metrics["total_books_rows"]}</div></div>', unsafe_allow_html=True)
        col1.caption(f"2B Portal Records: {metrics['total_gstr2b_rows']}")

        col2.markdown(f'<div class="metric-card"><div class="metric-label">Exact Matches</div><div class="metric-val metric-val-success">{metrics["exact_matches"]}</div></div>', unsafe_allow_html=True)
        col2.caption(f"Fuzzy Matches: {metrics['fuzzy_matches']}")

        col3.markdown(f'<div class="metric-card"><div class="metric-label">Claimable ITC</div><div class="metric-val metric-val-success">₹{metrics["total_itc_claimed"]:,.2f}</div></div>', unsafe_allow_html=True)
        col3.caption("Safe to claim in GSTR-3B")

        col4.markdown(f'<div class="metric-card"><div class="metric-label">Blocked / Trapped ITC</div><div class="metric-val metric-val-danger">₹{metrics["total_itc_blocked_risk"]:,.2f}</div></div>', unsafe_allow_html=True)
        col4.caption(f"{metrics['missing_in_2b']} vendor(s) not filed")

        col5.markdown(f'<div class="metric-card"><div class="metric-label">Mismatched Bills</div><div class="metric-val">{metrics["mismatches"]}</div></div>', unsafe_allow_html=True)
        col5.caption(f"Diff Tax: ₹{metrics['total_mismatch_itc']:,.2f}")

        # Carry-forward resolution alert if any
        if not resolved_cf.empty:
            st.info(f"🎉 **{len(resolved_cf)} invoice(s)** were resolved via Cross-Month Carry-Forward! Vendors filed late for previous periods.")
            with st.expander("View Resolved Late Filings"):
                st.dataframe(resolved_cf, use_container_width=True)

        # Tabbed Data Tables
        st.markdown("### Detailed Reconciliation Categories")
        t1, t2, t3, t4, t5, t6, t7 = st.tabs([
            f"✅ Exact Matches ({metrics['exact_matches']})",
            f"🟡 Fuzzy Matches ({metrics['fuzzy_matches']})",
            f"⚠️ Mismatches ({metrics['mismatches']})",
            f"🚨 Missing in GSTR-2B ({metrics['missing_in_2b']})",
            f"🔍 Missing in Books ({metrics['missing_in_books']})",
            f"🛡️ RCM Invoices ({metrics['rcm_count']})",
            "📋 All Records"
        ])

        cols_to_show = [
            "Raw_GSTIN", "Raw_Vendor_Name", "Raw_Invoice_Number", "Matched_2B_Invoice",
            "Raw_Invoice_Date", "Parsed_Taxable_Value", "Parsed_Total_Tax",
            "Match_Confidence", "Discrepancy_Note"
        ]

        with t1:
            st.dataframe(books[books["Match_Category"] == "Exact Match"][cols_to_show], use_container_width=True)

        with t2:
            st.dataframe(books[books["Match_Category"] == "Fuzzy Match"][cols_to_show], use_container_width=True)

        with t3:
            st.dataframe(books[books["Match_Category"] == "Mismatch"][cols_to_show], use_container_width=True)

        with t4:
            st.markdown("**These invoices are in your Purchase Book but your vendors have not uploaded them to the GST Portal.** Claiming ITC on these without filing may invite GST department scrutiny or notices.")
            st.dataframe(books[books["Match_Category"] == "Missing in GSTR-2B"][cols_to_show], use_container_width=True)
            
            # Button to save to SQLite carry-forward ledger
            if st.button("💾 Save Missing Invoices to Cross-Month Carry-Forward Ledger"):
                saved_count = db.save_missing_invoices(
                    books[books["Match_Category"] == "Missing in GSTR-2B"],
                    current_period
                )
                st.success(f"Successfully stored {saved_count} missing invoices for period '{current_period}' in local SQLite ledger.")

        with t5:
            st.markdown("**These invoices appear in your government GSTR-2B, but are not found in your internal Purchase Book.** Verify if you received these goods/services or if they represent unrecorded purchases.")
            cols_2b = ["Raw_GSTIN", "Raw_Vendor_Name", "Raw_Invoice_Number", "Raw_Invoice_Date", "Parsed_Taxable_Value", "Parsed_Total_Tax", "Discrepancy_Note"]
            st.dataframe(missing_b[cols_2b], use_container_width=True)

        with t6:
            st.markdown("**Reverse Charge Mechanism (RCM) Invoices:** Reverse charge tax must be paid in cash by the recipient before it can be claimed as ITC. They are kept isolated from regular B2B ITC matching.")
            st.dataframe(books[books["Match_Category"] == "RCM Invoice (Isolated)"][cols_to_show], use_container_width=True)

        with t7:
            st.dataframe(books, use_container_width=True)


# =============================================================================
# MODULE 3: REPORTS & DASHBOARD
# =============================================================================
elif nav_choice == "3. Reports & Dashboard":
    st.markdown('<div class="main-title">📈 Reports, Visuals & Vendor Follow-Up</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Download audit-ready Excel reports, view visual match analytics, and auto-draft vendor emails.</div>', unsafe_allow_html=True)

    with st.expander("💡 How this works", expanded=False):
        st.markdown("""
        - **Executive Visuals:** Visual breakdown of matched vs trapped tax credit and top vendors blocking your cash flow.
        - **Audit-Ready Excel:** Download a professionally styled, color-coded multi-tab workbook with a complete summary dashboard.
        - **Vendor Follow-up:** Instantly draft ready-to-send emails and WhatsApp reminders for vendors with missing invoices so they file GSTR-1 immediately.
        """)

    if st.session_state.get("reco_results") is None:
        st.info("No reconciliation has been run yet. Please go to tab **'2. Reconciliation'** to upload and reconcile files, or click **'Load Synthetic Test Data'** in the sidebar.")
    else:
        books = st.session_state["reco_results"]
        missing_b = st.session_state["missing_in_books"]
        metrics = st.session_state["reco_metrics"]

        # Visual Analytics
        st.markdown("### 📊 Visual Analytics")
        v_col1, v_col2 = st.columns(2)

        with v_col1:
            st.markdown("#### Invoice Count by Match Category")
            category_counts = books["Match_Category"].value_counts().reset_index()
            category_counts.columns = ["Category", "Count"]
            st.bar_chart(data=category_counts, x="Category", y="Count", color="#1E3A8A")

        with v_col2:
            st.markdown("#### Top Vendors by Blocked ITC (₹)")
            missing_2b = books[books["Match_Category"] == "Missing in GSTR-2B"]
            if not missing_2b.empty:
                vendor_blocked = (
                    missing_2b.groupby("Raw_Vendor_Name")["Parsed_Total_Tax"]
                    .sum()
                    .reset_index()
                    .rename(columns={"Raw_Vendor_Name": "Vendor", "Parsed_Total_Tax": "Blocked_ITC"})
                    .sort_values(by="Blocked_ITC", ascending=False)
                    .head(7)
                )
                st.bar_chart(data=vendor_blocked, x="Vendor", y="Blocked_ITC", color="#DC2626")
            else:
                st.success("No blocked ITC found across vendors!")

        st.divider()

        # Section 1: Excel Audit Report Download
        st.markdown("### 📥 Download Official Reconciliation Audit Pack")
        st.markdown("Export a complete, multi-tab Excel workbook with summary dashboard, color-coded status tabs, and notes.")

        excel_report_io = report_generator.generate_excel_reconciliation_report(books, missing_b, metrics)
        st.download_button(
            label="📊 Download Color-Coded Excel Reconciliation Report (.xlsx)",
            data=excel_report_io.getvalue(),
            file_name=f"GSTR2B_Reconciliation_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

        st.divider()

        # Section 2: Vendor Follow-Up Email & WhatsApp Generator
        st.markdown("### ✉️ Auto-Draft Vendor Follow-Up Communications")
        st.markdown("Quickly generate courteous, professional follow-up messages for vendors whose invoices are missing from your GSTR-2B.")

        missing_vendors_df = books[books["Match_Category"] == "Missing in GSTR-2B"]

        if missing_vendors_df.empty:
            st.success("🎉 All vendor invoices are present in GSTR-2B! No follow-up communications needed.")
        else:
            unique_vendors = missing_vendors_df["Raw_Vendor_Name"].dropna().unique().tolist()
            if not unique_vendors:
                unique_vendors = missing_vendors_df["Raw_GSTIN"].unique().tolist()

            sel_col1, sel_col2 = st.columns([2, 1])
            with sel_col1:
                selected_vendor = st.selectbox("Select Vendor to Draft Message for:", unique_vendors)
            with sel_col2:
                msg_format = st.radio("Message Format:", ["Email Draft", "WhatsApp Message"], horizontal=True)

            vendor_rows = missing_vendors_df[
                (missing_vendors_df["Raw_Vendor_Name"] == selected_vendor) |
                (missing_vendors_df["Raw_GSTIN"] == selected_vendor)
            ]

            vendor_gstin = vendor_rows["Raw_GSTIN"].iloc[0]
            invoices_info = []
            for _, r in vendor_rows.iterrows():
                invoices_info.append({
                    "invoice_no": r["Raw_Invoice_Number"],
                    "date": r["Raw_Invoice_Date"],
                    "taxable_value": r["Parsed_Taxable_Value"],
                    "total_tax": r["Parsed_Total_Tax"]
                })

            if msg_format == "Email Draft":
                email_content = report_generator.generate_vendor_email(
                    vendor_name=str(selected_vendor),
                    gstin=str(vendor_gstin),
                    invoices_list=invoices_info
                )
                st.text_area("Ready-to-Send Email Template (Editable):", value=email_content, height=320)
            else:
                whatsapp_content = report_generator.generate_vendor_whatsapp(
                    vendor_name=str(selected_vendor),
                    gstin=str(vendor_gstin),
                    inv_count=len(invoices_info),
                    total_tax=sum(i["total_tax"] for i in invoices_info)
                )
                st.text_area("Ready-to-Send WhatsApp Text:", value=whatsapp_content, height=140)

        st.divider()

        # Section 3: Cross-Month SQLite Ledger Viewer
        st.markdown("### 🗓️ Cross-Month Carry-Forward Ledger")
        st.markdown("Invoices carried forward from previous months that were saved to the local SQLite database.")
        ledger_df = db.get_all_pending_invoices()
        if ledger_df.empty:
            st.info("The carry-forward ledger is currently empty. To save missing invoices for future months, use the 'Save Missing Invoices' button in the Reconciliation tab.")
        else:
            st.dataframe(ledger_df, use_container_width=True)
            if st.button("Clear Carry-Forward Ledger"):
                db.clear_ledger()
                st.success("Ledger cleared.")
                st.rerun()
