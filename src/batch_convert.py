# src/batch_convert.py

import os
import re
import sys
from pathlib import Path
from datetime import datetime
from tqdm import tqdm
from main import detect_expected_sheets, convert_report

# === CONFIG ===
BASE_DIR = Path(r"C:\Users\jamie\Cantec Fire Alarms\Cantec Office - Documents\Cantec\Location Data")

# === DATE PATTERN ===
DATE_PATTERN = re.compile(r"([A-Za-z]+ \d{1,2}, \d{4})")

def parse_date_from_filename(filename: str):
    match = DATE_PATTERN.search(filename)
    if not match:
        return None
    try:
        return datetime.strptime(match.group(1), "%B %d, %Y")
    except ValueError:
        return None

def find_most_recent_v7_file(folder: Path):
    v7_files = [f for f in folder.glob("*.xlsx") if "V7" in f.name.upper()]
    dated_files = []

    for file in v7_files:
        file_date = parse_date_from_filename(file.name)
        if file_date:
            dated_files.append((file_date, file))

    if not dated_files:
        return None

    dated_files.sort(reverse=True, key=lambda tup: tup[0])
    return dated_files[0][1]

def get_deepest_folders_with_v7_files(base_dir: Path):
    count = 0
    for root, dirs, files in os.walk(base_dir):
        folder = Path(root)
        v7_files = [f for f in files if "V7" in f.upper() and f.endswith(".xlsx")]
        if v7_files:
            count += 1
            print(f"\r🔍 Found folders with V7 files: {count}", end='', flush=True)
            yield folder
    print()

def batch_convert_all_reports(overwrite_autoconverted: bool = False):
    print(f"🔎 Scanning for V7 folders in {BASE_DIR}...")
    target_folders = list(get_deepest_folders_with_v7_files(BASE_DIR))

    if not target_folders:
        print("❌ No folders with V7 files found.")
        return

    print(f"\n🚀 Starting conversion of {len(target_folders)} folders...\n")

    for folder in tqdm(target_folders, desc="Converting", unit="folder"):
        file = find_most_recent_v7_file(folder)
        if not file:
            tqdm.write(f"⚠️ {folder.name}: No valid V7 file with date found. Skipping.")
            continue

        original_name = file.name
        original_name = file.name
        autoconverted_name = re.sub(r"(?i)V7", "V8 AutoConverted", original_name)
        autoconverted_path = file.with_name(autoconverted_name)

       
        
        # === Now skip if already converted ===
        if autoconverted_path.exists() and not overwrite_autoconverted:
            tqdm.write(f"🟡 {folder.name}: Already converted (V8 AutoConverted exists). Skipping.")
            continue


        v8_name = re.sub(r"(?i)V7", "V8", original_name)
        v8_path = file.with_name(v8_name)
        if v8_path.exists() and "AutoConverted" not in v8_path.name:
            tqdm.write(f"⛔ {folder.name}: Manual V8 already exists. Skipping.")
            continue

        found_sheets = detect_expected_sheets(str(file))
        if not found_sheets:
            tqdm.write(f"⚠️ {folder.name}: No expected sheets found. Skipping.")
            continue

        if "ULC - C2.1-2.12" not in found_sheets:
            tqdm.write(f"🚫 {folder.name}: Missing required sheet 'ULC - C2.1-2.12'. Cleaning up and skipping.")

            for f in folder.glob("*.xlsx"):
                if "V8 AutoConverted" in f.name:
                    try:
                        f.unlink()
                        tqdm.write(f"🧹 {folder.name}: Deleted leftover file: {f.name}")
                    except Exception as cleanup_error:
                        tqdm.write(f"⚠️ {folder.name}: Failed to delete {f.name} – {cleanup_error}")
            continue

        tqdm.write(f"📄 {folder.name}: Converting {file.name}")


        output_file = convert_report(
            input_filepath=str(file),
            sheets_to_convert=found_sheets,
            progress_callback=None,
            save_to_input_dir=True
        )

        if output_file:
            converted_path = Path(output_file)
            final_path = file.with_name(autoconverted_name)
            if converted_path.exists():
                converted_path.rename(final_path)
                tqdm.write(f"💾 {folder.name}: Saved {final_path.name}")
            else:
                tqdm.write(f"❌ {folder.name}: Conversion finished but file not found.")
        else:
            tqdm.write(f"❌ {folder.name}: Conversion failed.")

# === CLI Entrypoint ===
if __name__ == "__main__":
    overwrite_flag = "--overwrite" in sys.argv
    batch_convert_all_reports(overwrite_autoconverted=overwrite_flag)
