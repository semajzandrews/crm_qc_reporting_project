# CRM Quality Assurance Automation Suite

This repository contains a professional-grade automation suite for validating and comparing HubSpot vs. Salesforce report data. The system utilizes a 3-step pipeline to ensure 100% data integrity and machine-verified accuracy.

## System Prerequisites

Before running the suite, ensure your environment is configured correctly.

1.  **Python 3.10+**: Ensure Python is installed and added to your PATH.
2.  **Dependencies**: Install the required libraries using the provided manifest.
    ```bash
    pip install -r requirements.txt
    ```
3.  **Source Files**: Place your raw PDF reports in your `Downloads` folder (or equivalent).
    *   **HubSpot Reports**: Ensure they are in a folder (e.g., `HubSpot 2`).
    *   **Salesforce Reports**: Ensure they are in a folder (e.g., `Salesforce 2`).

---

## Execution Protocol (Step-by-Step)

Run the following scripts in exact numerical order. Do not skip steps.

### **Step 01: Synchronization & Triage**
**Command:**
```bash
python3 01_SYNC_TARGET_FOLDERS.py
```
**What it does:**
*   Opens a window for you to select your **HubSpot** and **Salesforce** source folders.
*   Scans for 1:1 matching report pairs based on the client name.
*   **Auto-Triage:** Automatically moves duplicate or mismatched files into `TRIAGE_COLLISIONS` and `TRIAGE_ORPHANS` folders to ensure only perfect pairs are processed.
*   **Output:** Generates `targets.json` (The Master Map).

---

### **Step 02: Visual Comparison Engine**
**Command:**
```bash
python3 02_EXECUTE_COMPARISON_ENGINE.py
```
**First Run (Calibration):**
If running on a new machine, you must map your screen coordinates first:
```bash
python3 02_EXECUTE_COMPARISON_ENGINE.py calibrate
```
*Follow the on-screen wizard to hover over buttons and press Enter.*

**What it does:**
*   Reads the `targets.json` map.
*   Automates your web browser to upload each pair to the Diffchecker engine.
*   Exports a visual **Split-View PDF** highlighting any differences.
*   **Automation Note:** Do not use your mouse/keyboard while this script is running.
*   **Output:** Saves comparison PDFs to `~/Downloads/QA_ANALYTICS_RESULTS/`.

---

### **Step 03: Analytical Reporting & Evidence**
**Command:**
```bash
python3 03_GENERATE_ANALYTICAL_REPORTS.py
```
**What it does:**
*   Performs a "Machine Precision" scan of the text data in every report.
*   Applies **Cascade Logic**: If data discrepancies are found, it invalidates the Summary Page.
*   **Output:** Generates two final artifacts in `~/Downloads/QA_ANALYTICS_RESULTS/`:
    1.  **`QA_ANALYTICS_REPORT_FINAL.xlsx`**: The Master Excel Grid (0 = Match, 1 = Mismatch).
    2.  **`QA_TECHNICAL_EVIDENCE.md`**: A row-by-row technical log explaining *why* every grade was assigned.

---

## Output Location
All final deliverables are automatically saved to:
📂 **`~/Downloads/QA_ANALYTICS_RESULTS/`**
