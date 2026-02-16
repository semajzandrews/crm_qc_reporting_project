import openpyxl
import os
import re
import json
import tkinter as tk
from tkinter import simpledialog, messagebox
from pdfminer.high_level import extract_text
from datetime import datetime

# Relative path configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TARGETS_FILE = os.path.join(BASE_DIR, "targets.json")

# Load Configuration from Script 01
if not os.path.exists(TARGETS_FILE):
    print(f"Error: {os.path.basename(TARGETS_FILE)} not found. Please run Script 01 first.")
    exit(1)

with open(TARGETS_FILE, "r") as f:
    config_meta = json.load(f)

HS_DIR = config_meta.get("hs_dir")
SF_DIR = config_meta.get("sf_dir")
TEMPLATE_PATH = config_meta.get("template_path")
TARGETS_DICT = config_meta.get("matches", {})

# Results are always localized to the Template directory
RESULTS_DIR = os.path.join(os.path.dirname(TEMPLATE_PATH), "QA_ANALYTICS_RESULTS")
OUTPUT_EXCEL = os.path.join(RESULTS_DIR, 'QA_ANALYTICS_REPORT_FINAL.xlsx')
LOG_FILE = os.path.join(RESULTS_DIR, "QA_TECHNICAL_EVIDENCE.md")

SHEET_NAME = 'QA Report Test Tracker'
TESTER_NAME = "Semaj Andrews"

if not os.path.exists(RESULTS_DIR):
    os.makedirs(RESULTS_DIR)

# Identification Anchors - REFINED FOR OVERFLOW AWARENESS
SECTIONS = [
    {'key': 'Summary Page', 'marker': 'Year Over Year Comparison of Calls', 'next_marker': 'Calls by Agency'},
    {'key': 'Site Page', 'marker': 'Calls by Agency', 'next_marker': 'Calls by Day of Week'},
    {'key': 'Day of Week', 'marker': 'Calls by Day of Week', 'next_marker': 'Calls by Hour of Day'},
    {'key': 'Hour of Day', 'marker': 'Calls by Hour of Day', 'next_marker': 'Calls by Outcome'},
    {'key': 'Outcome', 'marker': 'Calls by Outcome', 'next_marker': 'Calls by Diagnosis'},
    {'key': 'Diagnosis', 'marker': 'Calls by Diagnosis', 'next_marker': None}
]

def clean_text(text):
    """Normalize whitespace and remove redundant headers to handle multi-page wraps."""
    if not text: return ""
    # Collapse whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def get_section_text(text, config):
    """
    Slices PDF text based on greedy anchors to capture multi-page tables.
    Matches from the FIRST start_marker to the FIRST next_marker found AFTER the start.
    """
    start_marker = config['marker']
    end_marker = config['next_marker']
    
    try:
        # 1. Find the FIRST occurrence of the start marker (Case-Insensitive)
        start_match = re.search(re.escape(start_marker), text, re.IGNORECASE)
        if not start_match:
            return None
        
        start_idx = start_match.start()
        
        # 2. Find the FIRST occurrence of the NEXT marker starting AFTER the start index
        if end_marker:
            end_match = re.search(re.escape(end_marker), text[start_idx + len(start_marker):], re.IGNORECASE)
            if end_match:
                # The index is relative to the slice, so we add the offset
                end_idx = start_idx + len(start_marker) + end_match.start()
            else:
                # If next marker not found, take everything to the end
                end_idx = len(text)
        else:
            end_idx = len(text)
            
        return text[start_idx:end_idx].strip()
    except Exception:
        return None

def process_client_analysis(sheet, row_idx, col_map, client_name, hs_path, sf_path, client_lines):
    print(f"--- PROCESSING: {client_name} ---")
    client_lines.append(f"\n### Client: {client_name}")
    
    if not os.path.exists(hs_path) or not os.path.exists(sf_path):
        msg = "Error: File paths could not be verified."
        print(f"   [{msg}]")
        client_lines.append(f"- {msg}")
        return False

    # 1. Binary comparison check (Skip heavy parsing if identical)
    with open(hs_path, 'rb') as f1, open(sf_path, 'rb') as f2:
        if f1.read() == f2.read():
            print("   Status: Exact binary match identified.")
            client_lines.append("- **Overall Result: 0** (Verified Binary Match)")
            sheet.cell(row=row_idx, column=col_map.get('Tester', 3)).value = TESTER_NAME
            report_col = col_map.get('Report ') or col_map.get('Report', 4)
            sheet.cell(row=row_idx, column=report_col).value = client_name
            for sec in SECTIONS:
                col_idx = col_map.get(sec['key'])
                if col_idx: sheet.cell(row=row_idx, column=col_idx).value = 0
            res_col = col_map.get('Test Result')
            if res_col: sheet.cell(row=row_idx, column=res_col).value = 0
            return True

    # 2. Sectional content analysis
    try:
        text_hs = extract_text(hs_path)
        text_sf = extract_text(sf_path)
    except Exception as e:
        msg = f"Error: Data extraction failed - {e}"
        print(f"   [{msg}]")
        client_lines.append(f"- {msg}")
        return False

    any_failure = False
    sheet.cell(row=row_idx, column=col_map.get('Tester', 3)).value = TESTER_NAME
    report_col = col_map.get('Report ') or col_map.get('Report', 4)
    sheet.cell(row=row_idx, column=report_col).value = client_name

    for section in SECTIONS:
        hs_block = get_section_text(text_hs, section)
        sf_block = get_section_text(text_sf, section)
        
        # LOGIC UPGRADE: If a section is missing from one but present in the other
        if hs_block is None and sf_block is None:
            result, reason = 0, "Section missing in both sources (Acceptable)"
        elif hs_block is None or sf_block is None:
            result, reason = 1, "Section presence mismatch"
        else:
            # NORMALIZATION: Collapse all whitespace to handle formatting shifts
            clean_hs = clean_text(hs_block)
            clean_sf = clean_text(sf_block)
            
            # VALIDATION: Check if content is just the header or too short to be valid
            if len(clean_hs) < 15 or len(clean_sf) < 15:
                # If they both match headers but have no data, 0. If one has data and other doesn't, 1.
                result = 0 if clean_hs == clean_sf else 1
                reason = "Empty table match" if result == 0 else "Content volume mismatch"
            else:
                result = 0 if clean_hs == clean_sf else 1
                reason = "Data Match" if result == 0 else "Data Discrepancy Identified"
            
        if result == 1: any_failure = True
        client_lines.append(f"- **{section['key']}**: {result} ({reason})")
        
        col_idx = col_map.get(section['key'])
        if col_idx: sheet.cell(row=row_idx, column=col_idx).value = result
        
    # 3. Final Integrity Guard: Summary Page inherits sub-page failures
    summary_key = 'Summary Page'
    summary_col = col_map.get(summary_key)
    if any_failure and summary_col:
        print(f"   Status: Discrepancy detected in sub-sections. Marking {summary_key} as 1.")
        sheet.cell(row=row_idx, column=summary_col).value = 1
        # Update log line for summary if it was previously marked 0
        for i, line in enumerate(client_lines):
            if f"**{summary_key}**: 0" in line:
                client_lines[i] = f"- **{summary_key}**: 1 (Inferred mismatch due to sub-page discrepancies)"
                break
    
    test_result_col = col_map.get('Test Result')
    overall = 1 if any_failure else 0
    if test_result_col: sheet.cell(row=row_idx, column=test_result_col).value = overall
    client_lines.append(f"- **Analytical Verdict**: {overall}")
    return True

def generate_final_analytics():
    """Generate Excel report and text log with batching and physical verification."""
    root = tk.Tk()
    root.withdraw()
    
    if not os.path.exists(OUTPUT_EXCEL):
        print("   Excel report missing. Resetting entries...")
        for name in TARGETS_DICT: TARGETS_DICT[name]['status_excel'] = 'pending'
    else:
        wb_check = openpyxl.load_workbook(OUTPUT_EXCEL)
        sheet_check = wb_check[SHEET_NAME]
        header_row = 3
        report_col_idx = 4
        for idx, cell in enumerate(sheet_check[header_row]):
            if cell.value and "Report" in str(cell.value):
                report_col_idx = idx + 1
                break
        existing_agencies = {str(sheet_check.cell(row=r, column=report_col_idx).value).strip() 
                             for r in range(header_row + 1, sheet_check.max_row + 1) 
                             if sheet_check.cell(row=r, column=report_col_idx).value}
        for name in TARGETS_DICT:
            TARGETS_DICT[name]['status_excel'] = 'completed' if name in existing_agencies else 'pending'

    ready_targets = {k: v for k, v in TARGETS_DICT.items() 
                     if v.get('status_pdf') == 'completed' and v.get('status_excel', 'pending') == 'pending'}
    total_ready = len(ready_targets)
    
    if total_ready == 0:
        if any(v.get('status_pdf', 'pending') == 'pending' for v in TARGETS_DICT.values()):
            messagebox.showwarning("Prerequisite Not Met", "No files are ready for Excel reporting. Run Script 02 first.")
        else:
            messagebox.showinfo("Complete", "All files have been successfully processed!")
        return

    batch_size = simpledialog.askinteger("Batch Size", f"Files Ready for Excel: {total_ready}\n\nHow many entries?", 
                                       initialvalue=total_ready, minvalue=1, maxvalue=total_ready)
    if not batch_size: return

    target_excel = OUTPUT_EXCEL if os.path.exists(OUTPUT_EXCEL) else TEMPLATE_PATH
    wb = openpyxl.load_workbook(target_excel)
    sheet = wb[SHEET_NAME]
    header_row = 3
    col_map = {str(cell.value).strip(): idx + 1 for idx, cell in enumerate(sheet[header_row]) if cell.value}

    report_col_idx = col_map.get('Report ') or col_map.get('Report', 4)
    current_row = header_row + 1
    while sheet.cell(row=current_row, column=report_col_idx).value: current_row += 1

    processed_count = 0
    for client_name, paths in ready_targets.items():
        if processed_count >= batch_size: break
        client_lines = []
        if process_client_analysis(sheet, current_row, col_map, client_name, paths['hs'], paths['sf'], client_lines):
            TARGETS_DICT[client_name]['status_excel'] = 'completed'
            config_meta['matches'] = TARGETS_DICT
            with open(TARGETS_FILE, "w") as f: json.dump(config_meta, f, indent=4)
            wb.save(OUTPUT_EXCEL)
            with open(LOG_FILE, "a") as f: f.write("\n".join(client_lines) + "\n")
            print(f"   [FINALIZED] {client_name} (Row {current_row})")
            current_row += 1
            processed_count += 1

    messagebox.showinfo("Batch Complete", f"Excel Batch Complete!\n\nProcessed: {processed_count}")

if __name__ == "__main__":
    generate_final_analytics()
