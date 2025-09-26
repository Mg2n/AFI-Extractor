# AFI-Extractor (Automation Script)

📄 This is a **Python automation script** that extracts AFIs (Areas for Improvement) from **Word (.docx)** and **PDF** reports, then writes them into a single **Excel file**.

⚠️ **Note:**  
This is **not a general-purpose project**.  
The script is designed for one specific use case: parsing reports where AFIs, Classifications, Entities, Recommendations, and Processes follow a fixed format.  
It will not work correctly on arbitrary documents.

---

## ✨ What it does
- Scans all `.docx` and `.pdf` files in the same folder.
- Looks for AFI sections:
  - **AFI (Area for Improvement)**
  - **Classification** (e.g., Major / Other)
  - **Entity** (extracted from `(Classification – Entity)`)
  - **Recommendation** (supports multi-line items)
  - **EE/FA (Process name/number)** from headings like:
    - `Process 1.1 Something`
    - or `Value / Operational / Business`
- Writes results into `All_AFIs.xlsx` with headers:

