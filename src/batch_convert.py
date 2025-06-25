# src/batch_convert.py

import os
import re
import sys
import json
import requests
import mimetypes
from pathlib import Path
from datetime import datetime
from rapidfuzz import fuzz, process
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed
from main import detect_expected_sheets, convert_report


BASE_DIR = Path(r"C:\Users\jamie\Cantec Fire Alarms\Cantec Office - Documents\Cantec\Location Data")
DATE_PATTERN = re.compile(r"([A-Za-z]+) (\d{1,2})(?:-\d{1,2})?, (\d{4})")
PROGRESS_LOG_PATH = BASE_DIR / "converted_folders.json"
FAILED_CONVERSIONS_LOG_PATH = BASE_DIR / "failed_conversions.json"
FAILED_UPLOADS_LOG_PATH = BASE_DIR / "failed_uploads.json"

GLOBAL_LOCATION_CACHE = []
NORMALIZED_LOCATION_MAP = {}  # normalized_address_str -> location_dict

self_session = requests.Session()

# region ServiceTrade Uploads
def upload_to_service_trade(file_path: Path, folder: Path, failed_upload_log):
    log("\n")
    log(f"Uploading {file_path.name} to ServiceTrade...")

    address_guess = extract_address_from_filename(file_path.name)
    log(f"Address guess from filename: {address_guess}")
    best_match = find_best_location_match(address_guess, NORMALIZED_LOCATION_MAP)

    if not best_match:
        log(f"No good match found for: {address_guess}. Skipping upload.")
        failed_upload_log.add(f"[No Location Match]: {address_guess}")
        return

    try:
        loc_id = best_match["location_id"]
        asset_id = best_match.get("asset_id")
        log(f"Match found: {best_match['full_text']} (Location ID: {loc_id}, Asset ID: {asset_id or 'N/A'})")

        if not asset_id:
            log(f"Skipping upload: No 'Building' asset found for location {loc_id}")
            failed_upload_log.add(f"[No Building Asset]: {address_guess}")
            return
        
        success = upload_file_to_servicetrade(file_path, asset_id)
        if not success:
            log(f"⚠️ Upload failed for {file_path.name}")
            failed_upload_log.add(f"[Upload Failed]: {address_guess}")


    except Exception as e:
        log(f"Upload failed for {file_path.name}: {e}")
        failed_upload_log.add(f"[Upload Failed with Exception {e}]: {address_guess}")
    log("\n")
    

def upload_file_to_servicetrade(file_path: Path, asset_id: int, description: str = ""):
    """
    Uploads a file to a ServiceTrade asset (entityType=3).
    PurposeId 3 = Job Picture (safest default).
    """
    endpoint = "https://api.servicetrade.com/api/attachment"
    purpose_id = 7  # "Generic Attachment"
    entity_type = 2  # 2 = Asset
    entity_id = asset_id

    file_path = Path(file_path)
    mime_type, _ = mimetypes.guess_type(file_path.name)
    if not mime_type:
        mime_type = "application/octet-stream"

    try:
        with open(file_path, "rb") as f:
            files = {
                "purposeId": (None, str(purpose_id)),
                "entityType": (None, str(entity_type)),
                "entityId": (None, str(entity_id)),
                "description": (None, description or f"Auto-uploaded report: {file_path.name}"),
                "uploadedFile": (file_path.name, f, mime_type),
            }

            response = self_session.post(endpoint, files=files)
            response.raise_for_status()

            json_data = response.json()
            uri = json_data.get("data", {}).get("uri")
            log(f"📎 Upload complete: {file_path.name} → {uri}")
            return True

    except Exception as e:
        log(f"❌ Failed to upload {file_path.name} – {e}")
        return False


def find_best_location_match(address_guess: str, location_map, threshold: int = 90):
    normalized_guess = normalize_address(address_guess)
    match, score, _ = process.extractOne(
        normalized_guess, location_map.keys(), scorer=fuzz.token_sort_ratio
    )

    if score >= threshold:
        return location_map[match]
    else:
        return None

def extract_address_from_filename(filename: str) -> str:
    try:
        name_without_ext = filename.rsplit(".", 1)[0]
        parts = name_without_ext.split(",")
        return parts[0].strip() if parts else name_without_ext
    except Exception as e:
        log(f"⚠️ Failed to extract address from filename '{filename}': {e}")
        return filename

def normalize_address(text: str) -> str:
    text = text.lower()
    text = re.sub(r'[^\w\s]', '', text)  # remove punctuation
    text = re.sub(r'\s+', ' ', text)  # collapse whitespace
    return text.strip()

def fetch_active_locations_from_st():
    """
    Fetch active locations from ServiceTrade API.
    Return a dictionary of location_id -> { address, name, full_text, asset_id, asset_name }
    """
    # Authenticate
    auth_url = "https://api.servicetrade.com/api/auth"
    payload = {"username": "jsullivan-phillips", "password": "Cetnac123!"}
    try:
        auth_response = self_session.post(auth_url, json=payload)
        auth_response.raise_for_status()
    except Exception as e:
        log(f"❌ Authentication failed: {e}")
        return {}

    # Grab total page count
    location_endpoint = "https://api.servicetrade.com/api/location"
    try:
        initial_response = self_session.get(location_endpoint, params={"status": "active", "limit": 200, "page": 1})
        initial_response.raise_for_status()
    except Exception as e:
        log(f"❌ Initial location retrieval failed: {e}")
        return {}

    total_pages = initial_response.json().get("data", {}).get("totalPages", 1)

    # Step 1: Multi-threaded fetch of all location pages
    def fetch_location_page(page_num):
        try:
            resp = self_session.get(location_endpoint, params={"status": "active", "limit": 200, "page": page_num})
            resp.raise_for_status()
            return resp.json().get("data", {}).get("locations", [])
        except Exception as e:
            log(f"❌ Failed to fetch page {page_num}: {e}")
            return []

    all_locations = []
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(fetch_location_page, p) for p in range(1, total_pages + 1)]
        for f in tqdm(as_completed(futures), total=len(futures), desc="📡 Fetching Location Pages"):
            all_locations.extend(f.result())

    # Step 2: Multi-threaded fetch of assets for each location
    active_locations = {}

    def fetch_asset_for_location(location):
        loc_id = location["id"]
        address = location.get("address", {}).get("street", "UNKNOWN")
        name = location.get("name", "")
        full_text = f"{name} {address} {location.get('address', {}).get('city', '')} {location.get('address', {}).get('state', '')} {location.get('address', {}).get('postalCode', '')}"

        result = {
            "location_id": loc_id,
            "name": name,
            "address": address,
            "full_text": full_text
        }

        asset_endpoint = "https://api.servicetrade.com/api/asset"
        asset_params = {"locationId": loc_id, "name": "Building"}

        try:
            asset_resp = self_session.get(asset_endpoint, params=asset_params)
            asset_resp.raise_for_status()
            asset_data = asset_resp.json().get("data", {}).get("assets", [])
            for asset in asset_data:
                if asset["name"] == "Building":
                    result.update({
                        "asset_id": asset["id"],
                        "asset_name": asset["name"]
                    })
                    break
        except Exception as e:
            log(f"❌ Asset fetch failed for location {loc_id} ({address}): {e}")

        return loc_id, result

    with ThreadPoolExecutor(max_workers=15) as executor:
        futures = [executor.submit(fetch_asset_for_location, loc) for loc in all_locations]
        for f in tqdm(as_completed(futures), total=len(futures), desc="🏢 Fetching Assets"):
            loc_id, data = f.result()
            active_locations[loc_id] = data

    log(f"✅ Retrieved {len(active_locations)} active locations with assets.")
    return active_locations

# endregion

def parse_date_from_filename(filename: str):
    match = DATE_PATTERN.search(filename)
    if not match:
        return None
    try:
        month_str, day_str, year_str = match.groups()
        date_str = f"{month_str} {day_str}, {year_str}"
        return datetime.strptime(date_str, "%B %d, %Y")
    except ValueError:
        return None

def find_most_recent_v7_file(folder: Path):
    v7_files = [
        f for f in folder.glob("*.xlsx")
        if "V7" in f.name.upper() and "sprinkler only" not in f.name.lower()
    ]
    dated_files = []
    for file in v7_files:
        file_date = parse_date_from_filename(file.name)
        if file_date:
            dated_files.append((file_date, file))
    if not dated_files:
        return None
    dated_files.sort(reverse=True, key=lambda tup: (tup[0].year, tup[0].month, tup[0].day))
    return dated_files[0][1]

def get_deepest_folders_with_v7_files(base_dir: Path):
    for root, dirs, files in os.walk(base_dir):
        folder = Path(root)
        v7_files = [f for f in files if "V7" in f.upper() and f.endswith(".xlsx")]
        if v7_files:
            yield folder

def load_progress():
    if PROGRESS_LOG_PATH.exists():
        try:
            with open(PROGRESS_LOG_PATH, "r") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def load_failed_upload_log():
    if FAILED_UPLOADS_LOG_PATH.exists():
        try:
            with open(FAILED_UPLOADS_LOG_PATH, "r") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def log(msg):
    tqdm.write(msg)
    sys.stdout.flush()

def load_failed_log():
    if FAILED_CONVERSIONS_LOG_PATH.exists():
        try:
            with open(FAILED_CONVERSIONS_LOG_PATH, "r") as f:
                return set(json.load(f))
        except Exception:
            return set()
    return set()

def save_failed_upload_log(failed_folders):
    with open(FAILED_UPLOADS_LOG_PATH, "w") as f:
        json.dump(sorted(failed_folders), f, indent=2)

def save_failed_log(failed_folders):
    with open(FAILED_CONVERSIONS_LOG_PATH, "w") as f:
        json.dump(sorted(failed_folders), f, indent=2)

def save_progress(processed_folders):
    with open(PROGRESS_LOG_PATH, "w") as f:
        json.dump(sorted(processed_folders), f, indent=2)

def batch_convert_all_reports(overwrite_autoconverted: bool = False):
    failed_log = load_failed_log()
    processed = load_progress()
    failed_upload_log = load_failed_upload_log()
    print("✅ Preview of processed entries:", list(processed)[:5])

    global GLOBAL_LOCATION_CACHE, NORMALIZED_LOCATION_MAP
    log("📡 Fetching active ServiceTrade locations...")
    GLOBAL_LOCATION_CACHE = fetch_active_locations_from_st()
    # Build normalized map once
    NORMALIZED_LOCATION_MAP = {
        normalize_address(loc_data["address"]): loc_data
        for loc_data in GLOBAL_LOCATION_CACHE.values()
    }

    log(f"🔎 Scanning for V7 folders in {BASE_DIR}...")
    target_folders = list(get_deepest_folders_with_v7_files(BASE_DIR))
    if not target_folders:
        log("❌ No folders with V7 files found.")
        return

   
    log(f"Found {len(processed)} folders already processed.\n")
    log(f"Starting conversion of {len(target_folders)} folders...\n")
    
    for folder in tqdm(target_folders, desc="Converting", unit="folder"):
        if str(folder) in processed:
            log(f"{folder.name}: Already processed. Skipping.")
            continue
        v8_uploaded = False
        try:
            file = find_most_recent_v7_file(folder)
            if not file:
                log(f"⚠️ {folder.name}: No valid V7 file with date found. Skipping.")
                processed.add(str(folder))
                save_progress(processed)
                continue
            else:
                original_name = file.name
                autoconverted_name = re.sub(r"(?i)V7", "V8 AutoConverted", original_name)
                autoconverted_path = file.with_name(autoconverted_name)

                if autoconverted_path.exists() and not overwrite_autoconverted:
                    log(f"🟡 {file.name}: Already converted. Skipping.")
                else:
                    v8_name = re.sub(r"(?i)V7", "V8", original_name)
                    v8_path = file.with_name(v8_name)
                    if v8_path.exists() and "AutoConverted" not in v8_path.name:
                        log(f"{file.name}: Manual V8 already exists. Skipping.")
                    else:
                        found_sheets = detect_expected_sheets(str(file))

                        if not found_sheets:
                            log(f"⚠️ {file.name}: No expected sheets found. Skipping.")
                        else:
                            # ❗ Check for ULC+C2.1 sheets
                            ulc_c2_sheets = [
                                s for s in found_sheets
                                if "ULC" in s.upper() and "C2.1" in s.upper()
                            ]
                            if len(ulc_c2_sheets) > 1:
                                log(f"{file.name}: Multiple ULC + C2.1 sheets found ({ulc_c2_sheets}). Skipping.")
                                failed_log.add(str(folder))
                                if autoconverted_path.exists():
                                    try:
                                        autoconverted_path.unlink()
                                        log(f"{file.name}: Deleted pre-existing AutoConverted file due to conflict.")
                                    except Exception as cleanup_error:
                                        log(f"{file.name}: Failed to delete {autoconverted_path.name} – {cleanup_error}")
                                continue


                            # ❗ Check for multiple "LOG REPORT" sheets
                            log_sheets = [s for s in found_sheets if "LOG REPORT" in s.upper()]
                            if len(log_sheets) > 1:
                                log(f"🔶 {file.name}: Multiple LOG REPORT sheets found ({log_sheets}). Skipping.")
                                failed_log.add(str(file))
                                if autoconverted_path.exists():
                                    try:
                                        autoconverted_path.unlink()
                                        log(f"{file.name}: Deleted pre-existing AutoConverted file due to conflict.")
                                    except Exception as cleanup_error:
                                        log(f"{file.name}: Failed to delete {autoconverted_path.name} – {cleanup_error}")
                                continue


                            if "ULC - C2.1-2.12" not in found_sheets:
                                log(f"{file.name}: Missing required sheet. Cleaning up and skipping.")
                                for f in folder.glob("*.xlsx"):
                                    if "V8 AutoConverted" in f.name:
                                        try:
                                            f.unlink()
                                            log(f"{file.name}: Deleted leftover file: {f.name}")
                                        except Exception as cleanup_error:
                                            log(f"{file.name}: Failed to delete {f.name} – {cleanup_error}")
                            else:
                                log(f"📄 {file.name}: Converting {file.name}")
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
                                        if final_path.exists():
                                            try:
                                                final_path.unlink()
                                            except Exception as e:
                                                log(f"{file.name}: Failed to delete existing {final_path.name} – {e}")
                                                continue
                                        converted_path.rename(final_path)
                                        log(f"{file.name}: Saved {final_path.name}")

                                        # 📤 Upload the successfully converted file to ServiceTrade
                                        upload_to_service_trade(final_path, folder, failed_upload_log)
                                        save_failed_upload_log(failed_upload_log)
                                        v8_uploaded = True
                                        
                                    else:
                                        log(f"{folder.name}: Conversion finished but file not found.")
                                        failed_log.add(f"[FAILED CONVERSION] {str(file)}")
                                else:
                                    log(f"{folder.name}: Conversion failed.")
                                    failed_log.add(f"[FAILED CONVERSION] {str(file)}")

        except Exception as e:
            name = file.name if file else folder.name
            log(f"🔥 {name}: Unexpected error – {e}")
            failed_log.add(f"[FAILED CONVERSION] {str(folder)}")
        
        finally:
            if not v8_uploaded and file and file.exists():
                log(f"📤 {file.name}: Uploading fallback original V7 to ServiceTrade...")
                upload_to_service_trade(file, folder, failed_upload_log)
                save_failed_upload_log(failed_upload_log)

        processed.add(str(folder))
        save_progress(processed)
        save_failed_log(failed_log)
        save_failed_upload_log(failed_upload_log)

if __name__ == "__main__":
    overwrite_flag = "--overwrite" in sys.argv
    batch_convert_all_reports(overwrite_autoconverted=overwrite_flag)
