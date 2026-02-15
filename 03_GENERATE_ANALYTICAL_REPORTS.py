import openpyxl
import os
import re
import json
from pdfminer.high_level import extract_text
from datetime import datetime

# Relative path configuration for portability
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS_DIR = os.path.expanduser("~/Downloads/")
HS_DIR = os.path.join(DOWNLOADS_DIR, "HubSpot 2/")
SF_DIR = os.path.join(DOWNLOADS_DIR, "Salesforce 2/")
RESULTS_DIR = os.path.join(DOWNLOADS_DIR, "QA_ANALYTICS_RESULTS/")
TEMPLATE_PATH = os.path.join(DOWNLOADS_DIR, 'Qlik Month End Reporting - Quality Assaunce Testing - Semaj.xlsx')
OUTPUT_EXCEL = os.path.join(RESULTS_DIR, 'QA_ANALYTICS_REPORT_FINAL.xlsx')
LOG_FILE = os.path.join(RESULTS_DIR, "QA_TECHNICAL_EVIDENCE.md")
TARGETS_FILE = os.path.join(BASE_DIR, "targets.json")

SHEET_NAME = 'QA Report Test Tracker'
TESTER_NAME = "Semaj Andrews"

# Ensure Results Directory Exists
if not os.path.exists(RESULTS_DIR):
    os.makedirs(RESULTS_DIR)

# Identification Anchors
SECTIONS = [
    {'key': 'Summary Page', 'marker': 'Year Over Year Comparison of Calls', 'next_marker': 'Calls by Agency'},
    {'key': 'Site Page', 'marker': 'Calls by Agency', 'next_marker': 'Calls by Day of Week'},
    {'key': 'Day of Week', 'marker': 'Calls by Day of Week', 'next_marker': 'Calls by Hour of Day'},
    {'key': 'Hour of Day', 'marker': 'Calls by Hour of Day', 'next_marker': 'Calls by Outcome'},
    {'key': 'Outcome', 'marker': 'Calls by Outcome', 'next_marker': 'Calls by Diagnosis'},
    {'key': 'Diagnosis', 'marker': 'Calls by Diagnosis', 'next_marker': None}
]

def get_section_text(text, config):
    """Slices PDF text based on markers using case-insensitive matching."""
    start_marker = config['marker']
    end_marker = config['next_marker']
    
    try:
        start_match = re.search(re.escape(start_marker), text, re.IGNORECASE)
        if not start_match:
            return None
        
        start_idx = start_match.start()
        end_idx = len(text)
        
        if end_marker:
            end_match = re.search(re.escape(end_marker), text[start_idx + len(start_marker):], re.IGNORECASE)
            if end_match:
                end_idx = start_idx + len(start_marker) + end_match.start()
                
        return text[start_idx:end_idx].strip()
    except Exception:
        return None

def process_client_analysis(sheet, row_idx, col_map, client_name, hs_file, sf_file, client_lines):
    """Parses PDF content and evaluates discrepancies across specific report sections."""
    print(f"--- PROCESSING: {client_name} ---")
    client_lines.append(f"\n### Client: {client_name}")
    hs_path = os.path.join(HS_DIR, hs_file)
    sf_path = os.path.join(SF_DIR, sf_file)
    
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
            result = 0
            reason = "Section missing in both sources (Acceptable)"
        elif hs_block is None or sf_block is None:
            result = 1
            reason = "Section presence mismatch"
        else:
            clean_hs = re.sub(r'\s+', ' ', hs_block).strip()
            clean_sf = re.sub(r'\s+', ' ', sf_block).strip()
            result = 0 if clean_hs == clean_sf else 1
            reason = "Data Match" if result == 0 else "Data Discrepancy Identified"
            
        if result == 1: any_failure = True
        client_lines.append(f"- **{section['key']}**: {result} ({reason})")
        
        col_idx = col_map.get(section['key'])
        if col_idx: sheet.cell(row=row_idx, column=col_idx).value = result
        
    # 3. Final conditional check for generated images
    if any_failure:
        summary_col = col_map.get('Summary Page')
        if summary_col:
            is_present_match = False
            for line in client_lines:
                if line.startswith("- **Summary Page**: 0 (Data Match)"):
                    is_present_match = True
                    break
            
            if is_present_match:
                print(f"   Status: Cascading failure identified for Summary Page.")
                sheet.cell(row=row_idx, column=summary_col).value = 1
                for i, line in enumerate(client_lines):
                    if line.startswith("- **Summary Page**: 0 (Data Match)"):
                        client_lines[i] = "- **Summary Page**: 1 (Inferred mismatch: Supplemental data discrepancies found)"
                        break
    
    test_result_col = col_map.get('Test Result')
    overall = 1 if any_failure else 0
    if test_result_col: sheet.cell(row=row_idx, column=test_result_col).value = overall
    client_lines.append(f"- **Analytical Verdict**: {overall}")
    return True

def generate_final_analytics():
    """Main process for generating comprehensive analytical reports with persistence."""
    for report_path in [OUTPUT_EXCEL, LOG_FILE]:
        if os.path.exists(report_path):
            try:
                os.remove(report_path)
                print(f"Initializing clean environment: Removed {os.path.basename(report_path)}")
            except Exception:
                pass

    if not os.path.exists(TARGETS_FILE):
        print(f"Error: Required file {os.path.basename(TARGETS_FILE)} not found.")
        return
        
    # Process all targets found in configuration
    with open(TARGETS_FILE, "r") as f:
        targets_dict = json.load(f)
    
    targets = []
    for name, files in targets_dict.items():
        targets.append((name, os.path.basename(files['hs']), os.path.basename(files['sf'])))

    print(f"Loading data template: {os.path.basename(TEMPLATE_PATH)}")
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    sheet = wb[SHEET_NAME]
    header_row = 3
    col_map = {str(cell.value).strip(): idx + 1 for idx, cell in enumerate(sheet[header_row]) if cell.value}

    with open(LOG_FILE, "w") as f:
        f.write(f"# Analysis Evidence Log - {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")

    row = header_row + 1
    for name, hs_file, sf_file in targets:
        client_lines = []
        if process_client_analysis(sheet, row, col_map, name, hs_file, sf_file, client_lines):
            wb.save(OUTPUT_EXCEL)
            with open(LOG_FILE, "a") as f:
                f.write("\n".join(client_lines) + "\n")
            print(f"   [Data Persisted] Record finalized for {name}")
        row += 1

    print(f"\nAnalytical sequence complete.")
    print(f"Report: {OUTPUT_EXCEL}")
    print(f"Evidence: {LOG_FILE}")

if __name__ == "__main__":
    generate_final_analytics()
