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
    """Map screen coordinates for automation with live execution during setup."""
    print("\n--- LIVE CALIBRATION SETUP ---")
    print("This mode will execute buttons as you capture them to reach the Export screen.")
    print("Ensure the browser is open and visible.")
    
    if not os.path.exists(TARGETS_FILE):
        print("Error: targets.json not found. Run Script 01 first.")
        return
        
    with open(TARGETS_FILE, "r") as f:
        meta = json.load(f)
        matches = meta.get("matches", {})
        if not matches:
            print("Error: No file pairs found.")
            return
        first_key = list(matches.keys())[0]
        hs_test_path = matches[first_key]["hs"]
        sf_test_path = matches[first_key]["sf"]

    new_config = {}
    try:
        input("👉 Hover over: [COMPARISON_AREA] and press ENTER...")
        pos = pyautogui.position()
        new_config["COMPARISON_AREA"] = [pos.x, pos.y]
        pyautogui.click(pos.x, pos.y)
        
        input("👉 Hover over: [LEFT_BROWSE] and press ENTER...")
        pos = pyautogui.position()
        new_config["LEFT_BROWSE"] = [pos.x, pos.y]
        
        input("👉 Hover over: [RIGHT_BROWSE] and press ENTER...")
        pos = pyautogui.position()
        new_config["RIGHT_BROWSE"] = [pos.x, pos.y]

        input("👉 Hover over: [FIND_DIFF_BTN] and press ENTER...")
        pos = pyautogui.position()
        new_config["FIND_DIFF_BTN"] = [pos.x, pos.y]

        print("\n🚀 Performing Live Upload to transition to Export Screen...")
        upload_sequence(new_config, hs_test_path, sf_test_path)
        
        print("\n⏳ Waiting 5 seconds for results to load...")
        time.sleep(5)

        input("👉 Hover over: [EXPORT_BTN] and press ENTER...")
        pos = pyautogui.position()
        new_config["EXPORT_BTN"] = [pos.x, pos.y]
        pyautogui.click(pos.x, pos.y)
        time.sleep(1.0)

        input("👉 Hover over: [SPLIT_VIEW_BTN] and press ENTER...")
        pos = pyautogui.position()
        new_config["SPLIT_VIEW_BTN"] = [pos.x, pos.y]
        pyautogui.click(pos.x, pos.y)
        time.sleep(2.0)

        input("👉 Hover over: [SAVE_BTN] and press ENTER...")
        pos = pyautogui.position()
        new_config["SAVE_BTN"] = [pos.x, pos.y]
        
        input("👉 Hover over: [TAB_CLOSE_BTN] and press ENTER...")
        pos = pyautogui.position()
        new_config["TAB_CLOSE_BTN"] = [pos.x, pos.y]

        input("👉 Hover over: [TAB_NEW_BTN] and press ENTER...")
        pos = pyautogui.position()
        new_config["TAB_NEW_BTN"] = [pos.x, pos.y]

        input("👉 Hover over: [DOCUMENT_MODE_BTN] and press ENTER...")
        pos = pyautogui.position()
        new_config["DOCUMENT_MODE_BTN"] = [pos.x, pos.y]

        with open(CONFIG_FILE, "w") as f:
            json.dump(new_config, f, indent=4)
        print(f"\n✨ Setup Complete!")
    except KeyboardInterrupt:
        print("\nAborted.")

def run_comparison_process(file_pairs):
    """Main execution loop for the comparison process."""
    if not os.path.exists(CONFIG_FILE):
        print(f"Error: Configuration file '{CONFIG_FILE}' not found.")
        return
    with open(CONFIG_FILE, "r") as f:
        coords = json.load(f)

    is_mac = platform.system() == "Darwin"
    print("\n--- Step 2: Running Comparisons ---")
    time.sleep(3)

    for i, (name, files) in enumerate(file_pairs.items()):
        print(f"\n[{i+1}/{len(file_pairs)}] File: {name}")
        upload_sequence(coords, files["hs"], files["sf"])
        time.sleep(6) 
        pyautogui.click(coords["EXPORT_BTN"])
        time.sleep(1.5)
        pyautogui.click(coords["SPLIT_VIEW_BTN"])
        time.sleep(2.5) 
        timestamp = datetime.now().strftime("%m%d_%H%M")
        base_name = f"{name}_Comparison_{timestamp}"
        
        # Removed Cmd+Shift+G to keep focus on default highlight
        pyautogui.write(base_name)
        time.sleep(1)
        pyautogui.press('enter')
        pyautogui.click(coords["SAVE_BTN"])
        time.sleep(4)

        # Move to Results
        expected_file = os.path.join(DOWNLOADS_DIR, base_name + ".pdf")
        if os.path.exists(expected_file):
            os.rename(expected_file, os.path.join(RESULTS_DIR, base_name + ".pdf"))

        pyautogui.click(coords["TAB_CLOSE_BTN"])
        time.sleep(1)
        pyautogui.click(coords["TAB_NEW_BTN"])
        time.sleep(2)
        pyautogui.click(coords["DOCUMENT_MODE_BTN"])
        time.sleep(2)

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "calibrate":
        calibrate_mode()
    else:
        with open(TARGETS_FILE, "r") as f:
            meta = json.load(f)
            targets = meta.get("matches", {})
        run_comparison_process(targets)
