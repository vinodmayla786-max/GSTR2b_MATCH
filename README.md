# GSTR2B Match: Indian GST Input Tax Credit (ITC) Reconciler & Tally Extractor

A production-ready Streamlit web application tailored for Indian MSMEs, tax practitioners, freelance accountants, and Chartered Accountant (CA) firms to:
1. **Reconcile GSTR-2B vs. Purchase Register** using multi-factor fuzzy matching with configurable tolerances.
2. **Extract Messy GST PDFs** (GSTR-2B, GSTR-1, GSTR-3B) with merged cells into structured Excel, CSV, and **Tally-ready XML**.
3. **Cross-Month Carry-Forward Tracking** with local SQLite database to automatically match invoices when vendors file late.
4. **Auto-Generate Vendor Follow-Up Emails and WhatsApp Alerts** to unblock trapped Input Tax Credit (ITC).

---

## 🚀 Quick Start (Local Setup for Non-Coders)

### 1. Install Python
- Download Python (version 3.10, 3.11, or 3.12) from [python.org](https://www.python.org/downloads/).
- **Important:** Check the box that says **"Add Python to PATH"** during installation.

### 2. Download and Extract the Code
- Place all files (`app.py`, `matcher.py`, `pdf_extractor.py`, `report_generator.py`, `db.py`, `sample_data_generator.py`, `requirements.txt`) into a single folder on your computer (e.g., `C:\GSTR2B_Match`).

### 3. Open Terminal or Command Prompt
- On Windows: Press `Win + R`, type `cmd`, and press Enter.
- Navigate to your folder:
  ```bash
  cd C:\GSTR2B_Match
  ```

### 4. Install Dependencies
Run this command once:
```bash
pip install -r requirements.txt
```

### 5. Launch the Web Application
Run:
```bash
streamlit run app.py
```
Your browser will automatically open with the application running at `http://localhost:8501`!

---

## ☁️ Free Cloud Deployment (Streamlit Community Cloud)

You can host this application online for free so your team or clients can use it anywhere:

1. **Upload to GitHub:**
   - Create a free account at [github.com](https://github.com).
   - Create a new repository named `gstr2b-match`.
   - Upload the files (`app.py`, helper modules, `requirements.txt`, and `README.md`).

2. **Deploy on Streamlit Community Cloud:**
   - Go to [share.streamlit.io](https://share.streamlit.io/) and log in with your GitHub account.
   - Click **"New app"**.
   - Select your repository (`gstr2b-match`), branch (`main`), and Main file path (`app.py`).
   - Click **"Deploy!"**.
   - Within 2 minutes, your web application will be live at a public URL (e.g., `https://your-firm-gstr2b.streamlit.app`).

---

## 📖 Module Guide

### Tab 1: PDF Extractor
- **Upload:** Any official GST return PDF (GSTR-2B, GSTR-1, or GSTR-3B) directly downloaded from the GST portal.
- **Auto-Fixes:** Automatically forward-fills supplier GSTIN and Trade Name for multi-line or merged-cell tables.
- **Downloads:**
  - **Clean Excel / CSV:** Standard flat table with columns for GSTIN, Name, Inv No, Date, Taxable Value, IGST, CGST, SGST, Cess.
  - **Tally XML Voucher:** Ready to import into TallyPrime or Tally.ERP 9 via **Gateway of Tally → Import Data → Vouchers**.
  - **Direct Bridge:** Click **"Send to Reconciliation Engine"** to seamlessly feed the extracted data into Tab 2 without re-uploading!

### Tab 2: Reconciliation
- **Upload:**
  1. Internal Purchase Register (`.csv` or `.xlsx`).
  2. GSTR-2B Portal Report (`.json`, `.xlsx`, `.csv`, or `.pdf`).
- **Matching Pipeline:**
  - Standardizes invoice numbers (removes `/`, `-`, spaces, and leading zeros).
  - Matches GSTIN + Invoice + Date + Amounts with configurable tolerances (e.g. ±₹1 or ±1%).
  - Uses fuzzy matching (`rapidfuzz`/`thefuzz`) to catch invoice number typos.
  - Isolates **RCM (Reverse Charge Mechanism)** invoices so they don't distort regular ITC.
- **Categories:**
  - **Exact Match:** Ready for 100% ITC claim.
  - **Fuzzy Match:** Matched with minor formatting discrepancies.
  - **Mismatch:** Both exist but amount or date differs beyond tolerance.
  - **Missing in GSTR-2B:** In your books, but vendor hasn't filed GSTR-1 (**Blocks ITC**).
  - **Missing in Purchase Book:** On portal, but unrecorded internally.

### Tab 3: Reports & Dashboard
- **Executive Visuals:** Chart of match distribution and top vendors with trapped ITC.
- **Audit Excel Pack:** Download an audit-ready multi-tab Excel workbook with color-coded sheets and a Summary Dashboard.
- **Vendor Follow-Up:** Select any vendor with missing invoices to auto-generate a polite, ready-to-send email or WhatsApp message with exact invoice details and blocked amounts.
- **Cross-Month Carry-Forward:** View and manage invoices saved in SQLite across multiple monthly tax cycles.
