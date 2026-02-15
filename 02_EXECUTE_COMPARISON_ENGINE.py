import pyautogui
import time
import json
import os
import sys
import platform
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

def generate_fallback_pdf(hs_path, sf_path, output_name):
    """Generates a Side-by-Side PDF locally when comparison engine fails to export."""
    print(f"   Generating local comparison report: {output_name}")
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
        full_output_path = os.path.join(DOWNLOADS_DIR, output_name)
        with open(full_output_path, "wb") as f:
            writer.write(f)
        print("   File generated successfully.")
    except Exception as e:
        print(f"   Generation failed: {e}")

def upload_sequence(coords, hs_path, sf_path):
    """Execute the file upload automation sequence with Cross-Platform support."""
    is_mac = platform.system() == "Darwin"
    
    if "COMPARISON_AREA" in coords:
        pyautogui.click(coords["COMPARISON_AREA"])
        time.sleep(0.5)
        
    # --- PRIMARY FILE ---
    print("   Uploading primary file...")
    pyautogui.click(coords["LEFT_BROWSE"])
    time.sleep(2.0) # Wait for dialog
    
    if is_mac:
        pyautogui.hotkey('command', 'shift', 'g')
        time.sleep(1.0)
    
    pyautogui.write(hs_path)
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)
    
    # --- SECONDARY FILE ---
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
    """Interactive wizard to map screen coordinates for the user."""
    print("\n--- CALIBRATION WIZARD ---")
    print("You will hover your mouse over specific buttons and press [ENTER] to capture coordinates.")
    print("Ensure your browser is open to Diffchecker and maximized.")
    print("Press [Ctrl+C] to abort at any time.\n")
    
    targets = [
        ("COMPARISON_AREA", "Click anywhere in the middle of the page to ensure focus"),
        ("LEFT_BROWSE", "Left 'Upload' or 'Browse' button"),
        ("RIGHT_BROWSE", "Right 'Upload' or 'Browse' button"),
        ("FIND_DIFF_BTN", "The main 'Find Difference' button"),
        ("EXPORT_BTN", "The 'Export' button (usually top right after diff)"),
        ("SPLIT_VIEW_BTN", "The 'Split View' option in the export menu"),
        ("SAVE_BTN", "The 'Save' or 'Download' button in the final dialog"),
        ("TAB_CLOSE_BTN", "The 'X' to close the current browser tab"),
        ("TAB_NEW_BTN", "The '+' to open a new browser tab"),
        ("DOCUMENT_MODE_BTN", "The bookmark or button to return to Diffchecker home")
    ]
    
    new_config = {}
    
    try:
        for key, desc in targets:
            input(f"👉 Hover over: [{desc}] and press ENTER...")
            pos = pyautogui.position()
            new_config[key] = [pos.x, pos.y]
            print(f"   ✅ Captured {key}: {pos}")
            time.sleep(0.5)
            
        with open(CONFIG_FILE, "w") as f:
            json.dump(new_config, f, indent=4)
            
        print(f"\n✨ Calibration Complete! Config saved to: {CONFIG_FILE}")
        
    except KeyboardInterrupt:
        print("\n\n❌ Calibration Aborted.")

def run_comparison_process(file_pairs):
    """Main execution loop for the comparison process."""
    if not os.path.exists(CONFIG_FILE):
        print(f"Error: Configuration file '{CONFIG_FILE}' not found.")
        return

    with open(CONFIG_FILE, "r") as f:
        coords = json.load(f)

    print("\n--- Phase 2: Comparison Engine Execution ---")
    print("Switch to browser interface in 3 seconds...")
    time.sleep(3)

    for i, (name, files) in enumerate(file_pairs.items()):
        print(f"\n[{i+1}/{len(file_pairs)}] PROCESSING: {name}")
        hs_path = files["hs"]
        sf_path = files["sf"]

        # 1. UI Automation Sequence
        upload_sequence(coords, hs_path, sf_path)
        time.sleep(5) 
        print("   Finalizing analysis...")
        pyautogui.click(coords["EXPORT_BTN"])
        time.sleep(1)
        pyautogui.click(coords["SPLIT_VIEW_BTN"])
        time.sleep(2) 
        timestamp = datetime.now().strftime("%m%d_%H%M")
        base_name = f"{name}_Comparison_{timestamp}"
        pyautogui.write(base_name)
        time.sleep(0.5)
        pyautogui.click(coords["SAVE_BTN"])
        time.sleep(3)

        # Verify Analysis Export and Move to Results
        expected_file = os.path.join(DOWNLOADS_DIR, base_name + ".pdf")
        final_destination = os.path.join(RESULTS_DIR, base_name + ".pdf")
        
        if os.path.exists(expected_file):
            os.rename(expected_file, final_destination)
            print(f"   ✅ Export Verified: {final_destination}")
        else:
            fallback_name = f"{base_name}_MATCH_REPORT.pdf"
            generate_fallback_pdf(hs_path, sf_path, fallback_name)
            output_fallback_path = os.path.join(DOWNLOADS_DIR, fallback_name)
            if os.path.exists(output_fallback_path):
                os.rename(output_fallback_path, os.path.join(RESULTS_DIR, fallback_name))

        # 3. Interface Reset
        print("   Resetting browser state...")
        pyautogui.click(coords["TAB_CLOSE_BTN"])
        time.sleep(1)
        pyautogui.click(coords["TAB_NEW_BTN"])
        time.sleep(2)
        pyautogui.click(coords["DOCUMENT_MODE_BTN"])
        time.sleep(2)

    print("Execution complete. Finalizing outputs.")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "calibrate":
        calibrate_mode()
    else:
        # Load targets from centralized configuration
        with open(TARGETS_FILE, "r") as f:
            targets = json.load(f)
        
        # Process all targets found in the configuration
        print(f"--- Processing {len(targets)} targets ---")
        run_comparison_process(targets)
