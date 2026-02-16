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
    """Map screen coordinates for automation."""
    print("\n--- SETUP ---")
    print("Hover over the specific buttons and press [ENTER] to save coordinates.")
    print("Ensure the browser is open and visible.")
    print("Press [Ctrl+C] to cancel.\n")
    
    targets = [
        ("COMPARISON_AREA", "Click anywhere in the middle of the page to ensure focus"),
        ("LEFT_BROWSE", "Left 'Upload' or 'Browse' button"),
        ("RIGHT_BROWSE", "Right 'Upload' or 'Browse' button"),
        ("FIND_DIFF_BTN", "The main 'Find Difference' button"),
        ("EXPORT_BTN", "The 'Export' button (usually top right after diff)"),
        ("SPLIT_VIEW_BTN", "The 'PDF - Side by Side' option in the export menu"),
        ("SAVE_BTN", "The BLUE 'Export' or 'Download' button in the final dialog (next to filename input)"),
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
            
        print(f"\n✨ Setup Complete! Settings saved to: {CONFIG_FILE}")
        
    except KeyboardInterrupt:
        print("\n\n❌ Calibration Aborted.")

def run_comparison_process(file_pairs):
    """Main execution loop for the comparison process."""
    if not os.path.exists(CONFIG_FILE):
        print(f"Error: Configuration file '{CONFIG_FILE}' not found.")
        return

    with open(CONFIG_FILE, "r") as f:
        coords = json.load(f)

    print("\n--- Step 2: Running Comparisons ---")
    print("Switch to browser in 3 seconds...")
    time.sleep(3)

    is_mac = platform.system() == "Darwin"

    for i, (name, files) in enumerate(file_pairs.items()):
        # Skip if already completed by Script 03
        if files.get('status') == 'completed':
            continue

        print(f"\nProcessing File: {name}")
        hs_path = files["hs"]
        sf_path = files["sf"]

        # 1. UI Automation Sequence
        upload_sequence(coords, hs_path, sf_path)
        
        # Long wait for Diffchecker to process the PDFs
        print("   Waiting for difference engine (10s)...")
        time.sleep(10) 
        
        print("   Opening Export menu...")
        pyautogui.click(coords["EXPORT_BTN"])
        time.sleep(2)
        
        print("   Selecting Side-by-Side PDF...")
        pyautogui.click(coords["SPLIT_VIEW_BTN"])
        time.sleep(3) # Wait for the filename/save dialog
        
        timestamp = datetime.now().strftime("%m%d_%H%M")
        base_name = f"{name}_Comparison_{timestamp}"
        
        print(f"   Entering filename: {base_name}")
        # On Mac, often need to ensure focus on the input or use hotkey
        if is_mac:
            pyautogui.hotkey('command', 'a')
            pyautogui.press('backspace')
            
        pyautogui.write(base_name)
        time.sleep(1)
        
        print("   Finalizing export...")
        pyautogui.click(coords["SAVE_BTN"])
        time.sleep(5) # Wait for download to finish

        # Verify Analysis Export and Move to Results
        expected_file = os.path.join(DOWNLOADS_DIR, base_name + ".pdf")
        final_destination = os.path.join(RESULTS_DIR, base_name + ".pdf")
        
        if os.path.exists(expected_file):
            os.rename(expected_file, final_destination)
            print(f"   ✅ Export Verified: {final_destination}")
        else:
            print(f"   ⚠️ Export check failed for {expected_file}. Using local fallback.")
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
    # Load targets from centralized configuration
    if not os.path.exists(TARGETS_FILE):
        print("Error: targets.json not found. Run Script 01 first.")
        sys.exit(1)

    with open(TARGETS_FILE, "r") as f:
        meta = json.load(f)
        targets = meta.get("matches", {})

    if len(sys.argv) > 1 and sys.argv[1] == "calibrate":
        calibrate_mode()
    else:
        # Process only pending targets if possible, or all if no status exists
        pending = {k: v for k, v in targets.items() if v.get('status') == 'pending'}
        
        if not pending:
            print("No pending comparisons found. (Check if you ran Script 01 or if all are 'completed')")
            # If nothing is 'pending', just process whatever is in matches
            pending = targets

        print(f"--- Processing {len(pending)} targets ---")
        run_comparison_process(pending)
