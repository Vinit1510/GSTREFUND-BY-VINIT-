# 📊 GST Rule 89(5) Refund Automation Tool - User Guide

Welcome to the **GST Refund Automation Tool**! This tool is designed to save you hours of manual work by instantly processing your GSTR-2B and GSTR-1 data to automatically calculate and generate the Annexure required for Inverted Duty Structure refunds under Rule 89(5).

Follow this simple guide to generate your perfect, ready-to-file PDF in minutes.

---

## 📥 Step 1: Download Required Files from the GST Portal

Before using the tool, you need to download your return data directly from the official GST Portal.

### 1. **GSTR-2B (Purchases / ITCs)**
- Log in to the GST Portal.
- Navigate to **Returns Dashboard**.
- Select the relevant financial year and month.
- Download the **GSTR-2B** in **EXCEL format** (`.xlsx`). 
- *(Note: Ensure you download the Excel file, not the JSON. The tool requires the standard Excel structure provided by the portal).*

### 2. **GSTR-1 (Sales)**
- Navigate to the **Returns Dashboard** for the relevant months.
- Download the **GSTR-1** in **JSON format**.
- **Important:** If you are claiming a refund for multiple months, download the JSON file for *every single month* in that period.

---

## 🚀 Step 2: Upload and Process Your Data

Open the web application and you will see three main tabs at the top of the screen.

### Tab 1: 📂 Upload GSTR-2B
1. Click on the **Upload GSTR-2B** tab.
2. Drag and drop the **GSTR-2B Excel file** you downloaded earlier into the upload box (or click to browse your computer).
3. The tool will instantly read the Excel sheets (`B2B`, `B2BA`, `CDNR`, `CDNRA`) and display a preview of your eligible Input Tax Credit (ITC) data.

### Tab 3: 📂 Upload GSTR-1
1. Click on the **Upload GSTR-1** tab.
2. Drag and drop **all of your GSTR-1 JSON files** into the upload box at the same time. 
3. *Yes! You can upload 3, 6, or even 12 months of JSON files all at once.*
4. Click the **"Process GSTR-1 Data"** button.
5. The tool will automatically extract, aggregate, and organize all sales data (`B2B`, `B2CS`, `B2CL`, `CDNR`) across all the months you uploaded.

---

## 🖨️ Step 3: Generate the Final PDF Report

Once your data is successfully uploaded and previewed, it's time to create your Annexure.

1. Click on the **📄 Generate Report** tab.
2. You will see a form asking for your business details. Fill in the following exactly as you want them to appear on the official report:
   - **Financial Year** (e.g., *2023-2024*)
   - **Tax Period** (e.g., *April 2023 to September 2023*)
   - **GSTIN** (Your 15-digit GST Number)
   - **Legal Name** (Your Business Name)
   - **Trade Name** (Optional)
3. Double-check your details, then click the **"Generate PDF"** button.

### Downloading Your Report
- The tool will instantly calculate the formulas, format the tables, and generate a strict, audit-ready **2-page PDF**.
- Click the **"Download Report"** button that appears.
- Your PDF will be perfectly scaled to A4 size with tight margins, ready to be printed, signed, and submitted to the GST Department!

---

### ⚠️ Troubleshooting & Tips
- **File Errors:** If you get an error when uploading, ensure you haven't accidentally opened and modified the files in Excel before uploading. Always use the raw, untouched files downloaded straight from the GST portal.
- **Missing Data:** Ensure your JSON files are the final filed versions of your GSTR-1.
- **Privacy:** Your data is 100% secure. This tool processes all calculations directly in the server memory and immediately deletes the files once you close the page. No financial data is ever saved or stored.
