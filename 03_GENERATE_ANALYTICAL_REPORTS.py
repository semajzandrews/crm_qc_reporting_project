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

SECTIONS = [
    {'key': 'Summary Page', 'marker': 'Year Over Year Comparison of Calls', 'next_marker': 'Calls by Agency'},
    {'key': 'Site Page', 'marker': 'Calls by Agency', 'next_marker': 'Calls by Day of Week'},
    {'key': 'Day of Week', 'marker': 'Calls by Day of Week', 'next_marker': 'Calls by Hour of Day'},
    {'key': 'Hour of Day', 'marker': 'Calls by Hour of Day', 'next_marker': 'Calls by Outcome'},
    {'key': 'Outcome', 'marker': 'Calls by Outcome', 'next_marker': 'Calls by Diagnosis'},
    {'key': 'Diagnosis', 'marker': 'Calls by Diagnosis', 'next_marker': None}
]

def clean_text(text):
    """Normalize whitespace to avoid spacing-related mismatches."""
    if not text: return ""
    return re.sub(r'\s+', ' ', text).strip()

def get_section_text(text, config):
    start_marker = config['marker']
    end_marker = config['next_marker']
    try:
        start_match = re.search(re.escape(start_marker), text, re.IGNORECASE)
        if not start_match: return None
        start_idx = start_match.start()
        end_idx = len(text)
        if end_marker:
            end_match = re.search(re.escape(end_marker), text[start_idx + len(start_marker):], re.IGNORECASE)
            if end_match:
                end_idx = start_idx + len(start_marker) + end_match.start()
        return text[start_idx:end_idx].strip()
    except Exception: return None

def process_client_analysis(sheet, row_idx, col_map, client_name, hs_path, sf_path, client_lines):
    print(f"--- PROCESSING: {client_name} ---")
    client_lines.append(f"\n### Client: {client_name}")
    
    if not os.path.exists(hs_path) or not os.path.exists(sf_path):
        msg = "Error: File paths could not be verified."
        print(f"   [{msg}]")
        client_lines.append(f"- {msg}")
        return False

    # 1. Binary comparison check
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
        
        if hs_block is None and sf_block is None:
            result, reason = 0, "Section missing in both sources (Acceptable)"
        elif hs_block is None or sf_block is None:
            result, reason = 1, "Section presence mismatch"
        else:
            result = 0 if clean_text(hs_block) == clean_text(sf_block) else 1
            reason = "Data Match" if result == 0 else "Data Discrepancy Identified"
            
        if result == 1: any_failure = True
        client_lines.append(f"- **{section['key']}**: {result} ({reason})")
        col_idx = col_map.get(section['key'])
        if col_idx: sheet.cell(row=row_idx, column=col_idx).value = result
        
    test_result_col = col_map.get('Test Result')
    overall = 1 if any_failure else 0
    if test_result_col: sheet.cell(row=row_idx, column=test_result_col).value = overall
    client_lines.append(f"- **Analytical Verdict**: {overall}")
    return True

def generate_final_analytics():
    """Generate Excel report and text log with batching."""
    root = tk.Tk()
    root.withdraw()
    
    pending_targets = {k: v for k, v in TARGETS_DICT.items() if v.get('status_excel', 'pending') == 'pending'}
    total_pending = len(pending_targets)
    
    if total_pending == 0:
        messagebox.showinfo("Complete", "No pending Excel analytical entries left!")
        return

    batch_size = simpledialog.askinteger("Batch Size", 
                                       f"Total Pending Excel Entries: {total_pending}\n\nHow many entries would you like to process?\n(Recommended: 10)", 
                                       initialvalue=10, minvalue=1, maxvalue=total_pending)
    if not batch_size: return

    # Load Template or existing result
    target_excel = OUTPUT_EXCEL if os.path.exists(OUTPUT_EXCEL) else TEMPLATE_PATH
    print(f"Opening: {os.path.basename(target_excel)}")
    wb = openpyxl.load_workbook(target_excel)
    sheet = wb[SHEET_NAME]
    header_row = 3
    col_map = {str(cell.value).strip(): idx + 1 for idx, cell in enumerate(sheet[header_row]) if cell.value}

    # Identify the next empty row for data persistence
    report_col_idx = col_map.get('Report ') or col_map.get('Report', 4)
    current_row = header_row + 1
    while sheet.cell(row=current_row, column=report_col_idx).value:
        current_row += 1

    processed_count = 0
    for client_name, paths in pending_targets.items():
        if processed_count >= batch_size: break
        
        client_lines = []
        if process_client_analysis(sheet, row_idx=current_row, col_map=col_map, client_name=client_name, 
                                   hs_path=paths['hs'], sf_path=paths['sf'], client_lines=client_lines):
            
            # Update status in config to prevent re-processing
            TARGETS_DICT[client_name]['status_excel'] = 'completed'
            config_meta['matches'] = TARGETS_DICT
            with open(TARGETS_FILE, "w") as f:
                json.dump(config_meta, f, indent=4)
                
            wb.save(OUTPUT_EXCEL)
            with open(LOG_FILE, "a") as f:
                f.write("\n".join(client_lines) + "\n")
            
            print(f"   [FINALIZED] {client_name} (Row {current_row})")
            current_row += 1
            processed_count += 1

    summary_msg = f"Excel Batch Complete!\n\nProcessed: {processed_count}\nRemaining: {total_pending - processed_count}"
    print(f"\n{summary_msg}")
    messagebox.showinfo("Batch Complete", summary_msg)

if __name__ == "__main__":
    generate_final_analytics()
