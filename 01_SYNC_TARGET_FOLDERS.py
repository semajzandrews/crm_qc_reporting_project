import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import shutil
import re

# Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TARGETS_FILE = os.path.join(BASE_DIR, "targets.json")

def select_directory(title):
    root = tk.Tk()
    root.withdraw()
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path

def synchronize_file_targets():
    """Multi-Pass matching strategy to maximize data integrity."""
    print("Initializing Multi-Pass Synchronization Engine...")
    
    hs_dir = select_directory("Select HubSpot Source Folder")
    if not hs_dir: return
    sf_dir = select_directory("Select Salesforce Source Folder")
    if not sf_dir: return

    # Results Destination
    RESULTS_MATCH_HS = os.path.join(hs_dir, "MATCHED_PAIRS")
    RESULTS_MATCH_SF = os.path.join(sf_dir, "MATCHED_PAIRS")
    for d in [RESULTS_MATCH_HS, RESULTS_MATCH_SF]:
        if not os.path.exists(d): os.makedirs(d)

    # Initial lists
    hs_pool = [f for f in os.listdir(hs_dir) if f.endswith('.pdf')]
    sf_pool = [f for f in os.listdir(sf_dir) if f.endswith('.pdf')]
    
    matches = {}

    # --- Phase 1: Exact Filename Match ---
    print("Phase 1: Checking for exact filename matches...")
    hs_pass1 = set(hs_pool)
    sf_pass1 = set(sf_pool)
    exact_matches = hs_pass1.intersection(sf_pass1)
    
    for filename in exact_matches:
        matches[filename] = {
            "hs": os.path.join(RESULTS_MATCH_HS, filename),
            "sf": os.path.join(RESULTS_MATCH_SF, filename)
        }
        shutil.move(os.path.join(hs_dir, filename), os.path.join(RESULTS_MATCH_HS, filename))
        shutil.move(os.path.join(sf_dir, filename), os.path.join(RESULTS_MATCH_SF, filename))
        hs_pool.remove(filename)
        sf_pool.remove(filename)

    # --- Phase 2: Timestamp Truncation Match ---
    # Example: 'Client_2026-02-10_14-30.pdf' vs 'Client_2026-02-11_15-00.pdf'
    print("Phase 2: Matching by truncating timestamps...")
    
    def get_base(f):
        return re.split(r'_202\d', f)[0].strip()

    hs_map = {get_base(f): f for f in hs_pool}
    sf_map = {get_base(f): f for f in sf_pool}
    
    trunc_bases = set(hs_map.keys()).intersection(set(sf_map.keys()))
    
    for base in trunc_bases:
        hs_filename = hs_map[base]
        sf_filename = sf_map[base]
        
        # Verify it's truly 1:1 at the truncated level to avoid collisions
        hs_count = sum(1 for f in hs_pool if get_base(f) == base)
        sf_count = sum(1 for f in sf_pool if get_base(f) == base)
        
        if hs_count == 1 and sf_count == 1:
            matches[base] = {
                "hs": os.path.join(RESULTS_MATCH_HS, hs_filename),
                "sf": os.path.join(RESULTS_MATCH_SF, sf_filename)
            }
            shutil.move(os.path.join(hs_dir, hs_filename), os.path.join(RESULTS_MATCH_HS, hs_filename))
            shutil.move(os.path.join(sf_dir, sf_filename), os.path.join(RESULTS_MATCH_SF, sf_filename))
            hs_pool.remove(hs_filename)
            sf_pool.remove(sf_filename)

    # --- Phase 3: Orphan Processing ---
    ORPHAN_HS = os.path.join(hs_dir, "UNMATCHED_ORPHANS")
    ORPHAN_SF = os.path.join(sf_dir, "UNMATCHED_ORPHANS")
    for d in [ORPHAN_HS, ORPHAN_SF]:
        if not os.path.exists(d): os.makedirs(d)
        
    for f in hs_pool: shutil.move(os.path.join(hs_dir, f), os.path.join(ORPHAN_HS, f))
    for f in sf_pool: shutil.move(os.path.join(sf_dir, f), os.path.join(ORPHAN_SF, f))

    with open(TARGETS_FILE, "w") as f:
        json.dump(matches, f, indent=4)
        
    summary = f"COMPLETED:\nMatches Found: {len(matches)}\nRemaining Orphans: {len(hs_pool) + len(sf_pool)}"
    print(f"\n{summary}")
    messagebox.showinfo("Extraction Complete", summary)

if __name__ == "__main__":
    synchronize_file_targets()
