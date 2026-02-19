import pyautogui
import time
import json
import os
import sys
import platform
import tkinter as tk
from tkinter import simpledialog, messagebox
from datetime import datetime
from pypdf import PdfReader, PdfWriter, PageObject, Transformation

# Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "diff_config.json")
TARGETS_FILE = os.path.join(BASE_DIR, "targets.json")
DOWNLOADS_DIR = os.path.expanduser("~/Downloads/")
RESULTS_DIR = os.path.join(DOWNLOADS_DIR, "QA_ANALYTICS_RESULTS/")

# Ensure Results Directory Exists
if not os.path.exists(RESULTS_DIR):
    os.makedirs(RESULTS_DIR)

def generate_side_by_side_pdf(hs_path, sf_path, output_name):
    """Generates a Side-by-Side merged PDF locally."""
    print(f"   Generating local report: {output_name}")
    try:
        reader_hs = PdfReader(hs_path)
        reader_sf = PdfReader(sf_path)
        writer = PdfWriter()
        num_pages = max(len(reader_hs.pages), len(reader_sf.pages))
        for i in range(num_pages):
            hs_width = 612.0
            hs_height = 792.0
            if i < len(reader_hs.pages):
                hs_page = reader_hs.pages[i]
                hs_width = float(hs_page.mediabox.width)
                hs_height = float(hs_page.mediabox.height)
            new_page = PageObject.create_blank_page(width=hs_width * 2, height=hs_height)
            if i < len(reader_hs.pages):
                new_page.merge_page(reader_hs.pages[i])
            if i < len(reader_sf.pages):
                sf_page = reader_sf.pages[i]
                op = Transformation().translate(tx=hs_width, ty=0)
                new_page.merge_transformed_page(sf_page, op)
            writer.add_page(new_page)
        full_output_path = os.path.join(RESULTS_DIR, output_name)
        with open(full_output_path, "wb") as f:
            writer.write(f)
        print(f"   ✅ Local Report Generated: {output_name}")
        return True
    except Exception as e:
        print(f"   ❌ Generation failed: {e}")
        return False

def upload_sequence(coords, hs_path, sf_path):
    """Execute the file upload automation sequence with Cross-Platform support."""
    is_mac = platform.system() == "Darwin"
    if "COMPARISON_AREA" in coords:
        pyautogui.click(coords["COMPARISON_AREA"])
        time.sleep(0.5)
    print("   Uploading primary file...")
    pyautogui.click(coords["LEFT_BROWSE"])
    time.sleep(2.0) 
    if is_mac:
        pyautogui.hotkey('command', 'shift', 'g')
        time.sleep(1.0)
    pyautogui.write(hs_path)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)
    print("   Uploading secondary file...")
    pyautogui.click(coords["RIGHT_BROWSE"])
    time.sleep(2.0)
    if is_mac:
        pyautogui.hotkey('command', 'shift', 'g')
        time.sleep(1.0)
    pyautogui.write(sf_path)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)
    print("   Initializing comparison engine...")
    pyautogui.click(coords["FIND_DIFF_BTN"])

def calibrate_mode():
    """Map screen coordinates for automation with live execution during setup."""
    print("\n--- LIVE CALIBRATION SETUP ---")
    if not os.path.exists(TARGETS_FILE):
        print("Error: targets.json not found.")
        return
    with open(TARGETS_FILE, "r") as f:
        meta = json.load(f)

    matches = meta.get("matches", {})
    if not matches:
        print("Error: No matched pairs found in targets.json.")
        return

    test_hs, test_sf, found_diff = "", "", False
    for name, paths in matches.items():
        try:
            with open(paths['hs'], 'rb') as f1, open(paths['sf'], 'rb') as f2:
                if f1.read() != f2.read():
                    test_hs, test_sf, found_diff = paths['hs'], paths['sf'], True
                    break
        except Exception: continue
    if not found_diff:
        first_key = list(matches.keys())[0]
        test_hs, test_sf = matches[first_key]["hs"], matches[first_key]["sf"]

    new_config = {}
    try:
        targets = [("COMPARISON_AREA", "Middle of page"), ("LEFT_BROWSE", "Left Browse"), 
                   ("RIGHT_BROWSE", "Right Browse"), ("FIND_DIFF_BTN", "Find Difference")]
        for key, desc in targets:
            input(f"👉 Hover over: [{desc}] and press ENTER...")
            pos = pyautogui.position()
            new_config[key] = [pos.x, pos.y]

        print("\n🚀 Transitioning to Export Screen...")
        upload_sequence(new_config, test_hs, test_sf)
        time.sleep(5)

        targets_v2 = [("EXPORT_BTN", "Top Right Export"), ("SPLIT_VIEW_BTN", "Side-by-Side PDF"), 
                      ("SAVE_BTN", "BLUE Export button"), ("TAB_CLOSE_BTN", "Tab Close 'X'"), 
                      ("TAB_NEW_BTN", "Tab New '+'"), ("DOCUMENT_MODE_BTN", "Diffchecker Home")]
        for key, desc in targets_v2:
            input(f"👉 Hover over: [{desc}] and press ENTER...")
            pos = pyautogui.position()
            new_config[key] = [pos.x, pos.y]
            if key in ["EXPORT_BTN", "SPLIT_VIEW_BTN"]: pyautogui.click(pos.x, pos.y); time.sleep(1)

        with open(CONFIG_FILE, "w") as f: json.dump(new_config, f, indent=4)
        print(f"\n✨ Setup Complete!")
    except KeyboardInterrupt: print("\nAborted.")

def run_comparison_process(config_meta):
    """Main execution loop with physical file verification."""
    if not os.path.exists(CONFIG_FILE): return
    with open(CONFIG_FILE, "r") as f: coords = json.load(f)
    
    targets_dict = config_meta.get("matches", {})
    
    # PHYSICAL VERIFICATION: If PDF is missing from RESULTS_DIR, set status_pdf back to 'pending'
    for name, files in targets_dict.items():
        if files.get('status_pdf') == 'completed':
            # We don't know the exact timestamp used previously, so we check for any PDF starting with the name
            pdf_found = any(f.startswith(name) and f.endswith(".pdf") for f in os.listdir(RESULTS_DIR))
            if not pdf_found:
                print(f"   Re-enabling PDF for {name} (File missing from Results)")
                targets_dict[name]['status_pdf'] = 'pending'

    pending_targets = {k: v for k, v in targets_dict.items() if v.get('status_pdf', 'pending') == 'pending'}
    total_pending = len(pending_targets)
    
    if total_pending == 0:
        messagebox.showinfo("Complete", "No pending PDF comparisons left!")
        return

    root = tk.Tk()
    root.withdraw()
    batch_size = simpledialog.askinteger("Batch Size", 
                                       f"Total Pending PDFs: {total_pending}\n\nHow many pairs would you like to process?\n(Recommended: 10)", 
                                       initialvalue=10, minvalue=1, maxvalue=total_pending)
    if not batch_size: return

    is_mac = platform.system() == "Darwin"
    print("\n--- Step 2: Running Comparisons ---")
    time.sleep(3)

    processed_count = 0
    for name, files in pending_targets.items():
        if processed_count >= batch_size: break
        
        print(f"\n[{processed_count+1}/{batch_size}] File: {name}")
        timestamp = datetime.now().strftime("%m%d_%H%M")
        
        is_identical = False
        with open(files["hs"], 'rb') as f1, open(files["sf"], 'rb') as f2:
            if f1.read() == f2.read(): is_identical = True

        if is_identical:
            print("   Status: Exact Match found. Generating Local Report...")
            out_name = f"{name}_MATCH_{timestamp}.pdf"
            generate_side_by_side_pdf(files["hs"], files["sf"], out_name)
        else:
            upload_sequence(coords, files["hs"], files["sf"])
            time.sleep(6) 
            pyautogui.click(coords["EXPORT_BTN"])
            time.sleep(1.5)
            pyautogui.click(coords["SPLIT_VIEW_BTN"])
            time.sleep(2.5) 
            
            base_name = f"{name}_Comparison_{timestamp}"
            pyautogui.write(base_name)
            time.sleep(2.0)
            pyautogui.click(coords["SAVE_BTN"])
            time.sleep(5)

            expected_file = os.path.join(DOWNLOADS_DIR, base_name + ".pdf")
            if os.path.exists(expected_file):
                os.rename(expected_file, os.path.join(RESULTS_DIR, base_name + ".pdf"))
            else:
                print("   ⚠️ Export failed. Generating Local Fallback...")
                out_name = f"{name}_Comparison_{timestamp}.pdf"
                generate_side_by_side_pdf(files["hs"], files["sf"], out_name)

            pyautogui.click(coords["TAB_CLOSE_BTN"]); time.sleep(1)
            pyautogui.click(coords["TAB_NEW_BTN"]); time.sleep(2)
            pyautogui.click(coords["DOCUMENT_MODE_BTN"]); time.sleep(2)

        # Update status in config
        targets_dict[name]['status_pdf'] = 'completed'
        config_meta['matches'] = targets_dict
        with open(TARGETS_FILE, "w") as f:
            json.dump(config_meta, f, indent=4)
        
        processed_count += 1

    summary_msg = f"PDF Batch Complete!\n\nProcessed: {processed_count}\nRemaining: {total_pending - processed_count}"
    print(f"\n{summary_msg}")
    messagebox.showinfo("Batch Complete", summary_msg)

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "calibrate":
        calibrate_mode()
    else:
        if not os.path.exists(TARGETS_FILE):
            print("Error: targets.json not found. Run Script 01 first.")
            sys.exit(1)
        with open(TARGETS_FILE, "r") as f:
            meta = json.load(f)
        run_comparison_process(meta)
