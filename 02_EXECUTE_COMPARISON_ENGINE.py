import pyautogui
import time
import json
import os
import sys
import platform

# Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "diff_config.json")
TARGETS_FILE = os.path.join(BASE_DIR, "targets.json")

def calibrate_mode():
    """Map screen coordinates for automation with live execution during setup."""
    print("\n--- LIVE CALIBRATION SETUP ---")
    print("This mode will execute buttons as you capture them to reach the Export screen.")
    print("Ensure the browser is open to the Diffchecker PDF upload page.")
    print("Press [Ctrl+C] to cancel.\n")
    
    # Check if we have targets to use for the live execution
    if not os.path.exists(TARGETS_FILE):
        print("Error: targets.json not found. Run Script 01 first to have files to test with.")
        return
        
    with open(TARGETS_FILE, "r") as f:
        meta = json.load(f)
        matches = meta.get("matches", {})
        if not matches:
            print("Error: No file pairs found in targets.json.")
            return
        # Get the first pair for the live test
        first_key = list(matches.keys())[0]
        hs_test_path = matches[first_key]["hs"]
        sf_test_path = matches[first_key]["sf"]

    is_mac = platform.system() == "Darwin"
    new_config = {}

    try:
        # 1. SETUP CLICKS
        input("👉 Hover over: [COMPARISON_AREA] (Middle of page) and press ENTER...")
        pos = pyautogui.position()
        new_config["COMPARISON_AREA"] = [pos.x, pos.y]
        pyautogui.click(pos.x, pos.y)
        
        # 2. LEFT BROWSE + LIVE UPLOAD
        input("👉 Hover over: [LEFT_BROWSE] (Left Upload button) and press ENTER...")
        pos = pyautogui.position()
        new_config["LEFT_BROWSE"] = [pos.x, pos.y]
        print(f"   Executing live upload for: {os.path.basename(hs_test_path)}")
        pyautogui.click(pos.x, pos.y)
        time.sleep(2.0)
        if is_mac:
            pyautogui.hotkey('command', 'shift', 'g')
            time.sleep(1.0)
        pyautogui.write(hs_test_path)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1.0)

        # 3. RIGHT BROWSE + LIVE UPLOAD
        input("👉 Hover over: [RIGHT_BROWSE] (Right Upload button) and press ENTER...")
        pos = pyautogui.position()
        new_config["RIGHT_BROWSE"] = [pos.x, pos.y]
        print(f"   Executing live upload for: {os.path.basename(sf_test_path)}")
        pyautogui.click(pos.x, pos.y)
        time.sleep(2.0)
        if is_mac:
            pyautogui.hotkey('command', 'shift', 'g')
            time.sleep(1.0)
        pyautogui.write(sf_test_path)
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1.0)

        # 4. FIND DIFFERENCE + LIVE EXECUTION
        input("👉 Hover over: [FIND_DIFF_BTN] (Find Difference) and press ENTER...")
        pos = pyautogui.position()
        new_config["FIND_DIFF_BTN"] = [pos.x, pos.y]
        print("   Executing Find Difference... Page will now transition.")
        pyautogui.click(pos.x, pos.y)
        
        print("\n⏳ Waiting 10 seconds for the Results screen to load...")
        time.sleep(10)

        # 5. NOW CAPTURE SECOND SCREEN BUTTONS
        input("👉 Page should be loaded. Hover over: [EXPORT_BTN] (Top Right) and press ENTER...")
        pos = pyautogui.position()
        new_config["EXPORT_BTN"] = [pos.x, pos.y]
        pyautogui.click(pos.x, pos.y)
        time.sleep(1.0)

        input("👉 Hover over: [SPLIT_VIEW_BTN] (PDF Side-by-Side) and press ENTER...")
        pos = pyautogui.position()
        new_config["SPLIT_VIEW_BTN"] = [pos.x, pos.y]
        pyautogui.click(pos.x, pos.y)
        time.sleep(2.0)

        input("👉 Hover over: [SAVE_BTN] (The BLUE Export/Save button) and press ENTER...")
        pos = pyautogui.position()
        new_config["SAVE_BTN"] = [pos.x, pos.y]
        
        # 6. RESET BUTTONS
        print("\nAlmost done. Return to the home page to capture reset buttons.")
        input("👉 Hover over: [TAB_CLOSE_BTN] (The 'X' on the browser tab) and press ENTER...")
        pos = pyautogui.position()
        new_config["TAB_CLOSE_BTN"] = [pos.x, pos.y]

        input("👉 Hover over: [TAB_NEW_BTN] (The '+' for a new tab) and press ENTER...")
        pos = pyautogui.position()
        new_config["TAB_NEW_BTN"] = [pos.x, pos.y]

        input("👉 Hover over: [DOCUMENT_MODE_BTN] (The bookmark/home button) and press ENTER...")
        pos = pyautogui.position()
        new_config["DOCUMENT_MODE_BTN"] = [pos.x, pos.y]

        with open(CONFIG_FILE, "w") as f:
            json.dump(new_config, f, indent=4)
            
        print(f"\n✨ Live Calibration Complete! Settings saved to: {CONFIG_FILE}")
        
    except KeyboardInterrupt:
        print("\n\n❌ Calibration Aborted.")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "calibrate":
        calibrate_mode()
    else:
        print("Use: python3 02_EXECUTE_COMPARISON_ENGINE.py calibrate")
