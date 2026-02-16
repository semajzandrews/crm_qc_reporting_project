# CRM Quality Assurance Automation Suite

A high-precision automation pipeline designed for the cross-platform validation and comparison of HubSpot and Salesforce report data. This suite implements a verified 3-step synchronization and analysis protocol.

## Environment Configuration

This project requires Python 3.10+ and should be executed within a virtual environment to ensure dependency stability.

1.  **Initialize Virtual Environment:**
    ```bash
    python3 -m venv venv
    ```
2.  **Activate Environment:**
    ```bash
    source venv/bin/activate
    ```
3.  **Install Dependencies:**
    ```bash
    pip3 install -r requirements.txt
    ```

---

## Execution Protocol

Scripts must be executed in numerical order. Ensure the virtual environment is active (`source venv/bin/activate`) before starting.

### **Step 01: Synchronization & Triage**
**Command:**
```bash
python3 01_SYNC_TARGET_FOLDERS.py
```
**Functionality:**
*   Initializes the source selection interface for HubSpot and Salesforce report directories.
*   **Automated Triage:** Filters duplicates and mismatches into `TRIAGE_COLLISIONS` and `TRIAGE_ORPHANS` directories.
*   **Output:** Generates `targets.json` (The Master Mapping File).

### **Step 02: Visual Comparison Engine**
**Command:**
```bash
python3 02_EXECUTE_COMPARISON_ENGINE.py
```

#### **📍 Calibration Protocol**
If executing on a new workstation or monitor configuration, initial coordinate mapping is required:
```bash
python3 02_EXECUTE_COMPARISON_ENGINE.py calibrate
```
**Instructions:**
1.  Open the web browser to the target comparison engine (Diffchecker).
2.  The CLI will request a hover-state for specific UI elements.
3.  Position the cursor over the target element; **do not click**.
4.  Press **Enter** to store the coordinate.
5.  Configuration is persistent in `diff_config.json`.

**Operational Note:** This script utilizes automated GUI interactions (PyAutoGUI). Do not use the mouse or keyboard while execution is in progress.

### **Step 03: Analytical Reporting & Evidence**
**Command:**
```bash
python3 03_GENERATE_ANALYTICAL_REPORTS.py
```
**Functionality:**
*   Performs machine-level OCR and text extraction on all comparison artifacts.
*   Applies validation logic to verify data integrity between reports.
*   **Output:** Generates `QA_ANALYTICS_REPORT_FINAL.xlsx` and `QA_TECHNICAL_EVIDENCE.md`.

---

## Output Directory
All deliverables are exported to:
**`~/Downloads/QA_ANALYTICS_RESULTS/`**

## Technical Note
Before beginning any session, verify that the `(venv)` indicator is present in your shell prompt. If not present, run:
```bash
source venv/bin/activate
```
